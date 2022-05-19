<%
Function BuildReport1471(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the number of checks by status
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1471"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim asQueries
	Dim asAreas
	Dim asCheckNumbers
	Dim sContents
	Dim lCurrentNumber
	Dim oRecordset
	Dim bEmpty
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim adTotals
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(sCondition, "Areas.", "Areas2."), "Companies.", "EmployeesHistoryListForPayroll."), "EmployeeTypes.", "EmployeesHistoryListForPayroll.")
	asQueries = Split(",,,", ",")
'	asQueries(0) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (Payments.StatusID In (-2,1,2)) And (CheckNumber Not In (Select Payments.ReplacementNumber As CheckNumber From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION />)) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
'	asQueries(1) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
'	asQueries(2) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.ReplacementNumber As CheckNumber, Payments.StatusID From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
'	asQueries(3) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (Payments.StatusID Not In (-2,1,2,3,4)) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
	asQueries(0) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (-2,1,2)) And (CheckNumber Not In ( Select Payments.ReplacementNumber As CheckNumber From  Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, EmployeeTypes, PayrollTypes, Banks, Areas As Areas1, Areas As Areas2 Where 1 = 1 And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION />)) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
	asQueries(1) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber" 
	asQueries(2) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.ReplacementNumber As CheckNumber, Payments.StatusID From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where 1 = 1 And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID=2) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber" 
	asQueries(3) = "Select BankName, Areas1.AreaID, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, PayrollTypeName, EmployeeTypeName, Payments.CheckNumber, Payments.StatusID From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where 1 = 1 And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID Not In (-2,1,2,3,4)) " & sCondition & " <CONDITION /> Order By Areas1.AreaCode, CheckNumber"
	sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Areas1.AreaID, Areas1.AreaCode From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') " & sCondition & " Order By Areas1.AreaCode", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'	Response.Write vbNewLine & "<!-- Query: Select Distinct Areas1.AreaID, Areas1.AreaCode From Payments, BankAccounts, Banks, Payrolls, PayrollTypes, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2 Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (Payments.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') " & sCondition & " Order By Areas1.AreaCode -->" & vbNewLine
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Areas1.AreaID, Areas1.AreaCode From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where 1 = 1 And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') " & sCondition & " Order By Areas1.AreaCode", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Distinct Areas1.AreaID, Areas1.AreaCode From Payments, Payrolls, EmployeesHistoryListForPayroll, EmployeesChangesLKP, PayrollTypes, EmployeeTypes, Banks, Areas As Areas1, Areas As Areas2 Where 1 = 1 And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.PaymentDate=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=PayrollTypes.PayrollTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (Areas2.ParentID=Areas1.AreaID) And (Payments.PaymentDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') " & sCondition & " Order By Areas1.AreaCode -->" & vbNewLine
	asAreas = ""
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asAreas = asAreas & CStr(oRecordset.Fields("AreaID").Value) & ","
			oRecordset.MoveNext
		Loop
	End If
	If Len(asAreas) > 0 Then
		asAreas = Left(asAreas, (Len(asAreas) - Len(",")))
		asAreas = Split(asAreas, ",")
		For kIndex = 0 To UBound(asAreas)
			bEmpty = True
			adTotals = Split("0,0,0,0", ",")
			For iIndex = 0 To UBound(asQueries)
				sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, Replace(asQueries(iIndex), "<CONDITION />", " And (Areas1.AreaID=" & asAreas(kIndex) & ")"), "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: " & Replace(asQueries(iIndex), "<CONDITION />", " And (Areas1.AreaID=" & asAreas(kIndex) & ")") & " -->" & vbNewLine
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						If bEmpty Then
							sContents = GetFileContents(Server.MapPath("Templates\Report_1471.htm"), sErrorDescription)
							sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
							sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayDateFromSerialNumber(lForPayrollID, -1, -1,- 1))
							sContents = Replace(sContents, "<BANK_NAME />", UCase(CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value))))
							sContents = Replace(sContents, "<AREA_NAME />", UCase(CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value) & ". " & CStr(oRecordset.Fields("AreaName1").Value))))
							sContents = Replace(sContents, "<PAYROLL_NUMBER />", GetPayrollNumber(lForPayrollID))
							sContents = Replace(sContents, "<PAYROLL_YEAR />", Left(lForPayrollID, Len("0000")))
							sContents = Replace(sContents, "<PAYROLL_TYPE />", UCase(CleanStringForHTML(CStr(oRecordset.Fields("PayrollTypeName").Value))))
							sContents = Replace(sContents, "<EMPLOYEE_TYPE />", UCase(CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value))))
						End If

						lCurrentNumber = -2
						asCheckNumbers = ""
						Do While Not oRecordset.EOF
							bEmpty = False
							If (CDbl(oRecordset.Fields("CheckNumber").Value) - lCurrentNumber) = 1 Then
							Else
								If lCurrentNumber > -2 Then asCheckNumbers = asCheckNumbers & "," & lCurrentNumber & LIST_SEPARATOR
								asCheckNumbers = asCheckNumbers & CDbl(oRecordset.Fields("CheckNumber").Value)
							End If
							lCurrentNumber = CDbl(oRecordset.Fields("CheckNumber").Value)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
						asCheckNumbers = asCheckNumbers & "," & lCurrentNumber
						asCheckNumbers = Split(asCheckNumbers, LIST_SEPARATOR)
						adTotals(iIndex) = 0
						For jIndex = 0 To UBound(asCheckNumbers)
							asCheckNumbers(jIndex) = Split(asCheckNumbers(jIndex), ",")
							sContents = Replace(sContents, "<CHECKS_0_" & iIndex & " />", asCheckNumbers(jIndex)(0) & "<BR /><CHECKS_0_" & iIndex & " />")
							sContents = Replace(sContents, "<CHECKS_1_" & iIndex & " />", asCheckNumbers(jIndex)(1) & "<BR /><CHECKS_1_" & iIndex & " />")
							sContents = Replace(sContents, "<CHECKS_2_" & iIndex & " />", (CDbl(asCheckNumbers(jIndex)(1)) - CDbl(asCheckNumbers(jIndex)(0)) + 1) & "<BR /><CHECKS_2_" & iIndex & " />")
							adTotals(iIndex) = adTotals(iIndex) + (CDbl(asCheckNumbers(jIndex)(1)) - CDbl(asCheckNumbers(jIndex)(0)) + 1)
						Next
						For jIndex = jIndex To 20
							sContents = Replace(sContents, "<CHECKS_0_" & iIndex & " />", "<BR /><CHECKS_0_" & iIndex & " />")
							sContents = Replace(sContents, "<CHECKS_1_" & iIndex & " />", "<BR /><CHECKS_1_" & iIndex & " />")
							sContents = Replace(sContents, "<CHECKS_2_" & iIndex & " />", "<BR /><CHECKS_2_" & iIndex & " />")
						Next
						sContents = Replace(sContents, "<TOTAL_CHECKS_" & iIndex & " />", adTotals(iIndex))
					End If
				End If
			Next
			If bEmpty Then
				lErrorNumber = -1
				sErrorDescription = "No existen cheques generados que cumplan con los criterios del filtro."
			Else
				For iIndex = 0 To UBound(adTotals)
					sContents = Replace(sContents, "<CHECKS_0_" & iIndex & " />", "&nbsp;")
					sContents = Replace(sContents, "<CHECKS_1_" & iIndex & " />", "&nbsp;")
					sContents = Replace(sContents, "<CHECKS_2_" & iIndex & " />", "&nbsp;")
					sContents = Replace(sContents, "<TOTAL_CHECKS_" & iIndex & " />", 0)
				Next
				sContents = Replace(sContents, "<TOTAL_CHECKS />", (adTotals(0) - adTotals(1) + adTotals(2) - adTotals(3)))
				Response.Write sContents
			End If
		Next
	End If

	Set oRecordset = Nothing
	BuildReport1471 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1472(oRequest, oADODBConnection, iPaymentTypeID, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the blocked deposits for the given payroll
'Inputs:  oRequest, oADODBConnection, iPaymentTypeID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1472"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim adTotal
	Dim iCurrentTypeID
	Dim lCurrentBankID
	Dim lCurrentAreaID
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim bPayrollIsClosed
    Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(sCondition, "Areas.", "Areas2."), "Companies.", "EmployeesHistoryList.")
	Select Case iPaymentTypeID
		Case -1
			Select Case oRequest("PaymentType").Item
				Case "0"
					sCondition = sCondition & " And (BankAccounts.AccountNumber='.')"
				Case "1"
					sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.')"
			End Select
			sCondition = sCondition & " And (Payments.StatusID Not In (-2,-1,1,2,3,4))"
		Case 0 'Pagos cancelados
			Select Case oRequest("PaymentType").Item
				Case "0"
					sCondition = sCondition & " And (BankAccounts.AccountNumber='.')"
				Case "1"
					sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.')"
			End Select
			sCondition = sCondition & " And (Payments.StatusID Not In (-2,-1,1,2,3,4))"
		Case 1 'Depósitos bloqueados
			sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.') And (Payments.StatusID In (4))"
	End Select
	If lPayrollID = -1 Then
	Else
		sCondition = sCondition & " And (PaymentDate=" & lPayrollID & ")"
	End If
	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)
	sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
	If bPayrollIsClosed Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypes.TypeID, EmployeesHistoryListForPayroll.BankID, BankName, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, PaymentCenters.AreaCode As PaymentCenterShortName, CheckAmount, Payments.Description From Payments, Banks, Employees, EmployeesHistoryListForPayroll, EmployeeTypes, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeeTypes.TypeID, BankName, Areas2.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        Response.Write vbNewLine & "<!-- Query: Select EmployeeTypes.TypeID, EmployeesHistoryListForPayroll.BankID, BankName, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, PaymentCenters.AreaCode As PaymentCenterShortName, CheckAmount, Payments.Description From Payments, Banks, Employees, EmployeesHistoryListForPayroll, EmployeeTypes, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.BankID=Banks.BankID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeeTypes.TypeID, BankName, Areas2.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypes.TypeID, Banks.BankID, BankName, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, PaymentCenters.AreaCode As PaymentCenterShortName, CheckAmount, Payments.Description From Payments, BankAccounts, Banks, Employees, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeeTypes.TypeID, BankName, Areas2.AreaCode, EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        Response.Write vbNewLine & "<!-- Query: Select EmployeeTypes.TypeID, Banks.BankID, BankName, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, PaymentCenters.AreaCode As PaymentCenterShortName, CheckAmount, Payments.Description From Payments, BankAccounts, Banks, Employees, EmployeesChangesLKP, EmployeesHistoryList, EmployeeTypes, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones, Areas As PaymentCenters Where (Payments.AccountID=BankAccounts.AccountID) And (BankAccounts.BankID=Banks.BankID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas2.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeeTypes.TypeID, BankName, Areas2.AreaCode, EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1472.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000"))))
			Response.Write sContents
			asColumnsTitles = Split("&nbsp;,No. del empleado,Nombre,Centro de pago, Importe,Observaciones", ",", -1, vbBinaryCompare)
			asCellAlignments = Split(",,,,RIGHT,", ",", -1, vbBinaryCompare)
			adTotal = Split(",", ",")
			adTotal(0) = Split("0,0", ",")
			adTotal(0)(0) = 0
			adTotal(0)(1) = 0
			adTotal(1) = Split("0,0", ",")
			adTotal(1)(0) = 0
			adTotal(1)(1) = 0
			iCurrentTypeID = -2
			lCurrentBankID = -2
			lCurrentAreaID = -2
			Do While Not oRecordset.EOF
				If (iCurrentTypeID <> CInt(oRecordset.Fields("TypeID").Value)) Or (lCurrentBankID <> CLng(oRecordset.Fields("BankID").Value)) Then
					If iCurrentTypeID <> -2 Then
						sRowContents = "<SPAN COLS=""3"" />&nbsp;" & TABLE_SEPARATOR & "<B>TOTAL EMPLEADOS:</B>" & TABLE_SEPARATOR & "<B>" & adTotal(0)(0) & "</B>" & TABLE_SEPARATOR
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						sRowContents = "<SPAN COLS=""3"" />&nbsp;" & TABLE_SEPARATOR & "<B>TOTAL IMPORTE:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0)(1), 2, True, False, True) & "</B>" & TABLE_SEPARATOR
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
						adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
						adTotal(0)(0) = 0
						adTotal(0)(1) = 0
						Response.Write "</TABLE><BR /><BR />"
					End If
					Response.Write "<CENTER><B>" & Replace(Replace(CStr(oRecordset.Fields("TypeID").Value), "0", "FUNCIONARIOS"), "1", "OPERATIVOS") & "<BR />REPORTE DE BLOQUEOS APLICADOS EN QNA: " & DisplayNumericDateFromSerialNumber(lPayrollID) & "</B><BR /></CENTER>"
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

					iCurrentTypeID = CInt(oRecordset.Fields("TypeID").Value)
					lCurrentBankID = -2
					lCurrentAreaID = -2
				End If
				If lCurrentBankID <> CLng(oRecordset.Fields("BankID").Value) Then
					sRowContents = "<SPAN COLS=""6"" /><B>BANCO: " & CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value)) & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					lCurrentBankID = CLng(oRecordset.Fields("BankID").Value)
				End If
				If lCurrentAreaID <> CLng(oRecordset.Fields("AreaID2").Value) Then
					sRowContents = "<SPAN COLS=""6"" /><B>ADSCRIPCIÓN: " & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value) & " " & CStr(oRecordset.Fields("AreaName2").Value)) & "</B>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					lCurrentAreaID = CLng(oRecordset.Fields("AreaID2").Value)
				End If

				If bForExport Then
					sRowContents = "&nbsp;" & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
				Else
					sRowContents = "&nbsp;" & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
				End If
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Description").Value))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adTotal(0)(0) = adTotal(0)(0) + 1
				adTotal(0)(1) = adTotal(0)(1) + CDbl(oRecordset.Fields("CheckAmount").Value)

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close

			sRowContents = "<SPAN COLS=""3"" />&nbsp;" & TABLE_SEPARATOR & "TOTAL EMPLEADOS:" & TABLE_SEPARATOR & "<B>" & adTotal(0)(0) & "</B>" & TABLE_SEPARATOR
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
			sRowContents = "<SPAN COLS=""3"" />&nbsp;" & TABLE_SEPARATOR & "<B>TOTAL IMPORTE:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(0)(1), 2, True, False, True) & "</B>" & TABLE_SEPARATOR
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
			adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
			adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
			Response.Write "</TABLE><BR /><BR />"
			Response.Write "<B>TOTAL GENERAL DE EMPLEADOS: " & adTotal(1)(0) & "</B><BR />"
			Response.Write "<B>TOTAL GENERAL LIQUIDO: " & FormatNumber(adTotal(1)(1), 2, True, False, True) & "</B><BR />"
		Else
			lErrorNumber = -1
			sErrorDescription = "No existen depósitos cancelados que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1472 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1474(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for bank deposits
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1474"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sPayrollName
	Dim bPayrollIsClosed
	Dim oRecordset
	Dim dTotal
	Dim lCounter
	Dim sRowContents
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim aTemp
	Dim sTemp
	Dim iIndex
	Dim lErrorNumber
	Dim asRecordSpei

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	sFilePath = Server.MapPath(sFilePath)
	Response.Flush()

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	lPayrollNumber = GetPayrollNumber(lForPayrollID)
	Call GetNameFromTable(oADODBConnection, "Payrolls", lPayrollID, "", "", sPayrollName, "")
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Areas.", "Areas2."), "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "-17", "17")

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	sErrorDescription = "No se pudieron obtener los depósitos bancarios."
	If bPayrollIsClosed Then
		If InStr(1,sCondition,"ParentZones.ZoneID",vbBinaryCompare) <> 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount As CheckAmount, EmployeesHistoryListForPayroll.BankID, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.","EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptAmount As CheckAmount, EmployeesHistoryListForPayroll.BankID, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount As CheckAmount, EmployeesHistoryListForPayroll.BankID, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.","EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptAmount As CheckAmount, EmployeesHistoryListForPayroll.BankID, EmployeesHistoryListForPayroll.AccountNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		End If
	Else
		If InStr(1,sCondition,"ParentZones.ZoneID",vbBinaryCompare) <> 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount As CheckAmount, BankAccounts.BankID, BankAccounts.AccountNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber<>'.') And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptAmount As CheckAmount, BankAccounts.BankID, BankAccounts.AccountNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Zones2.StartDate<=" & lForPayrollID & ") And (Zones2.EndDate>=" & lForPayrollID & ") And (ParentZones.StartDate<=" & lForPayrollID & ") And (ParentZones.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber<>'.') And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount As CheckAmount, BankAccounts.BankID, BankAccounts.AccountNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID = Employees.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber<>'.') And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptAmount As CheckAmount, BankAccounts.BankID, BankAccounts.AccountNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas2.AreaShortName, Areas2.AreaCode From Payroll_" & lPayrollID & ", BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber<>'.') And (BankAccounts.Active=1) And (EmployeesHistoryList.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & lPayrollID & ") And (StatusID Not In (-2,-1,1)))) And (Payroll_" & lPayrollID & ".ConceptID=0) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
		End If
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			dTotal = 0
			lCounter = 0
			Select Case CLng(oRecordset.Fields("BankID").Value)
				Case 1 'BBVA Bancomer
					lCounter = 1
					Select Case oRequest("BankID").Item
						Case "-1" 'BBVA Bancomer. Baja California
							Do While Not oRecordset.EOF
								'001001
								sRowContents = Right(("00000000000000000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("00000000000000000000")) 'Número del empleado
								sRowContents = sRowContents & "01" & Right(("0000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("0000000000")) 'Importe
								sRowContents = sRowContents & Right(lPayrollID, Len("YY")) & Mid(lPayrollID, Len("YYYYM"), Len("MM")) & Mid(lPayrollID, Len("YYY"), Len("YY")) 'Fecha de nómina
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                "), Len("                ")) 'Cuenta bancaria
								sRowContents = sRowContents & Right(lPayrollID, Len("YY")) & Mid(lPayrollID, Len("YYYYM"), Len("MM")) & Mid(lPayrollID, Len("YYY"), Len("YY")) 'Fecha de nómina
								sTemp = " "
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
								Err.Clear
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & "              "), Len("              ")) 'Nombre del empleado
								lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
						Case "-2" 'BBVA Bancomer. Pagel
							Do While Not oRecordset.EOF
								sRowContents = "3"
								sRowContents = Right(("00" & lCounter), Len("00")) 'Secuencia
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeNumber").Value) & "                    "), Len("                    ")) 'No. del empleado
								sRowContents = sRowContents & Right(("000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000")) 'Monto
								dTotal = dTotal + Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)
								sRowContents = Right(("000000000" & lCounter), Len("000000000")) 'Consecutivo
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("RFC").Value) & "               "), Len("               ")) 'RFC
								sRowContents = sRowContents & "98" 'Tarjeta de pagos
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                    "), Len("                    ")) 'Cuenta bancaria
								sRowContents = sRowContents & Right(lPayrollID, Len("YYMMDD"))
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                "), Len("                ")) 'Cuenta bancaria
								sRowContents = sRowContents & "              "
								sRowContents = sRowContents & "00"
								sRowContents = sRowContents & "    "

								lCounter = lCounter + 1
								lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
							sRowContents = "1"
							sRowContents = sRowContents & Right(("0000000" & lCounter), Len("0000000")) 'Registros
							sRowContents = sRowContents & Right(("000000000000000" & dTotal), Len("000000000000000")) 'Monto
							sRowContents = sRowContents & "0000000"
							sRowContents = sRowContents & "000000000000000"
							sRowContents = sRowContents & Right(("000000000000" & oRequest("Field05").Item), Len("000000000000")) 'No. de contrato
							sRowContents = sRowContents & "             "
							sRowContents = sRowContents & "101" 'Tipo de servicio
							sRowContents = sRowContents & "0"
							sRowContents = sRowContents & "      "
							sRowContents = sRowContents & vbNewLine
							sRowContents = sRowContents & GetFileContents(sFilePath, sErrorDescription)
							lErrorNumber = SaveTextToFile(sFilePath, sRowContents, sErrorDescription)
						Case Else 'BBVA Bancomer
							Do While Not oRecordset.EOF
								'001001
								sRowContents = Right(("000000000" & lCounter), Len("000000000")) 'Consecutivo
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("RFC").Value) & "               "), Len("               ")) 'RFC
								sRowContents = sRowContents & "98" 'Tarjeta de pagos
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                    "), Len("                    ")) 'Cuenta bancaria
								sRowContents = sRowContents & Right(("000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000")) 'Monto
								sTemp = " "
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
								Err.Clear
								sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeNumber").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & "                                        "), Len("                                        ")) 'Nombre del empleado
								sRowContents = sRowContents & "001" 'Banco destino
								sRowContents = sRowContents & "001" 'Plaza destino

								lCounter = lCounter + 1
								lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
					End Select
				Case 3 'Banamex
					If StrComp(oRequest("LayoutType").Item, "B", vbBinaryCompare) = 0 Then
						Do While Not oRecordset.EOF
							sRowContents = "30001"
							sRowContents = sRowContents & Right(("000000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000000")) 'Monto
							aTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value) & LIST_SEPARATOR, LIST_SEPARATOR)
							If Len(aTemp(1)) > 0 Then
								sRowContents = sRowContents & "01" 'Tipo de cuenta
							Else
								sRowContents = sRowContents & "03" 'Tipo de cuenta
							End If
							sRowContents = sRowContents & Right(("0000" & aTemp(1)), Len("0000")) 'No. de sucursal
							sRowContents = sRowContents & Right(("00000000000000000000" & aTemp(0)), Len("00000000000000000000")) 'No. de cuenta
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AreaCode").Value) & "0000000000"), Len("0000000000")) 'Referencia***
							sRowContents = sRowContents & "                                        " 'Referencia***
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
							Err.Clear
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeNumber").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & "                                        "), Len("                                                       ")) 'Beneficiario
							sRowContents = sRowContents & "                                        " 'Instrucciones
							sRowContents = sRowContents & "                        " 'Descripción
							sRowContents = sRowContents & "00" 'Clave de estado***
							sRowContents = sRowContents & "0000" 'Clave de ciudad***
							sRowContents = sRowContents & "0000" 'Clave de banco***
							dTotal = dTotal + Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)
							lCounter = lCounter + 1
							lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						sRowContents = "4001"
						sRowContents = sRowContents & Right(("000000" & lCounter), Len("000000"))
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						sRowContents = sRowContents & "000001"
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)

						sRowContents = "1" 'Tipo de registro
						sRowContents = sRowContents & Right(("000000000000" & oRequest("Field02").Item), Len("000000000000")) 'No. del cliente '000058364781|000059667242
						'If oRecordset.Fields("BankID").Value = 3 Then
                        '    Response.Write "Banamex"
                            sRowContents = sRowContents & oRequest("PayrollDepositDay").Item & oRequest("PayrollDepositMonth").Item & Right(oRequest("PayrollDepositYear").Item, Len("00")) 'Fecha de pago
                        'Else
                        '    sRowContents = sRowContents & Right(oRequest("PayrollIssueDay").Item & oRequest("PayrollIssueMonth").Item & oRequest("PayrollIssueYear").Item, Len("00")) 'Fecha de pago
                        'End If
                        sRowContents = sRowContents & Right(("0000" & oRequest("FileNumber").Item), Len("0000")) 'Secuencial
						sRowContents = sRowContents & "ISSSTE                              " 'Nombre de la empresa
						sRowContents = sRowContents & Left(sPayrollName, Len("                    ")) 'Descripción del archivo
						sRowContents = sRowContents & "05" 'Pagomático
						sRowContents = sRowContents & "                                        " 'Instrucciones
						sRowContents = sRowContents & "B01"
						sRowContents = sRowContents & vbNewLine

						sRowContents = sRowContents & "21001"
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						sRowContents = sRowContents & "01"
						sRowContents = sRowContents & Right(("0000" & oRequest("Field03").Item), Len("0000")) 'Sucursal '0100|0224
						sRowContents = sRowContents & Right(("00000000000000000000" & oRequest("Field04").Item), Len("00000000000000000000")) 'Cuenta de cargo "00000000000007708668|00000000000004160460"
						sRowContents = sRowContents & "                    "
						sRowContents = sRowContents & vbNewLine
						sRowContents = sRowContents & GetFileContents(sFilePath, sErrorDescription)
						lErrorNumber = SaveTextToFile(sFilePath, sRowContents, sErrorDescription)
					Else
						Do While Not oRecordset.EOF
							sRowContents = "30001"
							sRowContents = sRowContents & Right(("000000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000000")) 'Monto
							aTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value) & LIST_SEPARATOR, LIST_SEPARATOR)
							If Len(aTemp(1)) > 0 Then
								sRowContents = sRowContents & "01" 'Tipo de cuenta
							Else
								sRowContents = sRowContents & "03" 'Tipo de cuenta
							End If
							sRowContents = sRowContents & Right(("00000000000000000000" & aTemp(0)), Len("00000000000000000000"))
							'sRowContents = sRowContents & Right(("0000000000" & CStr(oRecordset.Fields("AreaShortName").Value)), Len("0000000000"))
							sRowContents = sRowContents & "                                        "
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
							Err.Clear
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeNumber").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & "                                        "), Len("                                                       "))
							sRowContents = sRowContents & "                                        "
							sRowContents = sRowContents & "                        "
							sRowContents = sRowContents & "    "
							sRowContents = sRowContents & "0000000"
							sRowContents = sRowContents & "00"
							dTotal = dTotal + Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)
							lCounter = lCounter + 1
							lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						sRowContents = "4001"
						sRowContents = sRowContents & Right(("000000" & lCounter), Len("000000"))
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						sRowContents = sRowContents & "000001"
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)

						sRowContents = "1"
						sRowContents = sRowContents & Right(("000000000000" & oRequest("Field02").Item), Len("000000000000")) 'No. del cliente '000058364781|000059667242
						sRowContents = sRowContents & oRequest("PayrollIssueDay").Item & oRequest("PayrollIssueMonth").Item & Right(oRequest("PayrollIssueYear").Item, Len("00"))
						sRowContents = sRowContents & Right(("0000" & oRequest("FileNumber").Item), Len("0000"))
						sRowContents = sRowContents & "ISSSTE                              "
						sRowContents = sRowContents & Left(sPayrollName, Len("                    "))
						sRowContents = sRowContents & "05"
						sRowContents = sRowContents & "                                        "
						sRowContents = sRowContents & "C01"
						sRowContents = sRowContents & vbNewLine

						sRowContents = sRowContents & "21001"
						sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000"))
						sRowContents = sRowContents & "01"
						sRowContents = sRowContents & Right(("0000" & oRequest("Field03").Item), Len("0000")) 'Sucursal '0100|0224
						sRowContents = sRowContents & Right(("00000000000000000000" & oRequest("Field04").Item), Len("00000000000000000000")) 'Cuenta de cargo "00000000000007708668|00000000000004160460"
						sRowContents = sRowContents & "                    "
						sRowContents = sRowContents & vbNewLine
						sRowContents = sRowContents & GetFileContents(sFilePath, sErrorDescription)
						lErrorNumber = SaveTextToFile(sFilePath, sRowContents, sErrorDescription)
					End If
				Case 14 'Banorte
					Do While Not oRecordset.EOF
						sRowContents = "D" 'Tipo de registro
						sRowContents = sRowContents & oRequest("PayrollIssueYear").Item & oRequest("PayrollIssueMonth").Item & oRequest("PayrollIssueDay").Item
						sRowContents = sRowContents & Right(("0000000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("0000000000")) 'Número del empleado
						sRowContents = sRowContents & "                                        " 'Referencia del servicio
						sRowContents = sRowContents & "                                        " 'Referencia leyenda del ordenante
						sRowContents = sRowContents & Right(("000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000")) 'Importe
						sRowContents = sRowContents & "072" 'Número del banco receptor
						sRowContents = sRowContents & "01" 'Tipo de cuenta
						sRowContents = sRowContents & Right(("000000000000000000" & CStr(oRecordset.Fields("AccountNumber").Value)), Len("000000000000000000")) 'Número de cuenta
						sRowContents = sRowContents & "0" 'Tipo de movimiento
						sRowContents = sRowContents & " " 'Acción
						sRowContents = sRowContents & "00000000" 'IVA
						sRowContents = sRowContents & "                  " 'Filler

						dTotal = dTotal + Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)
						lCounter = lCounter + 1
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					sRowContents = "H" 'Tipo de registro
					sRowContents = sRowContents & "NE" 'Clave de servicio
					sRowContents = sRowContents & Right(("00000" & oRequest("Field01").Item), Len("00000")) 'Emisora
					sRowContents = sRowContents & oRequest("PayrollIssueYear").Item & oRequest("PayrollIssueMonth").Item & oRequest("PayrollIssueDay").Item 'Fecha de proceso
					sRowContents = sRowContents & "01" 'Consecutivo
					sRowContents = sRowContents & Right(("000000" & lCounter), Len("000000")) 'Número de registros
					sRowContents = sRowContents & Right(("000000000000000" & dTotal), Len("000000000000000")) 'Importe total
					sRowContents = sRowContents & "000000" 'Número de altas
					sRowContents = sRowContents & "000000000000000" 'Importe de altas
					sRowContents = sRowContents & "000000" 'Número debajas
					sRowContents = sRowContents & "000000000000000" 'Importe de bajas
					sRowContents = sRowContents & "000000" 'Número de cuentas a verificar
					sRowContents = sRowContents & "0" 'Acción
					sRowContents = sRowContents & "                                                                             " 'Filler
					sRowContents = sRowContents & vbNewLine
					sRowContents = sRowContents & GetFileContents(sFilePath, sErrorDescription)
					lErrorNumber = SaveTextToFile(sFilePath, sRowContents, sErrorDescription)
				Case 17 'Serfín
					If StrComp(oRequest("BankID").Item, "-17", vbBinarycompare) = 0 Then 'Serfín. Honorarios
						Do While Not oRecordset.EOF
							sRowContents = "65501643330     " 'Cuenta de cargo
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                "), Len("                ")) 'Número de cuenta
							sRowContents = sRowContents & Right(("0000000000000" & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, False)), Len("0000000000000")) 'Importe
							sRowContents = sRowContents & "                                        "
							sRowContents = sRowContents & oRequest("ApplicationYear").Item & oRequest("ApplicationMonth").Item & oRequest("ApplicationDay").Item
							lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
					Else
						sRowContents = "1" 'Tipo de registro
						sRowContents = sRowContents & "00001" 'Número de secuencia
						sRowContents = sRowContents & "E" 'Sentido
						sRowContents = sRowContents & oRequest("FileMonth").Item & oRequest("FileDay").Item & oRequest("FileYear").Item 'Fecha de generación
						If InStr(1, oRequest("EmployeeTypeID").Item, "1", vbBinaryCompare) > 0 Then
							sRowContents = sRowContents & "65501643358     " 'Cuenta de cargo Funcionarios
						Else
							sRowContents = sRowContents & "65501643330     " 'Cuenta de cargo Operativos
						End If
						sRowContents = sRowContents & oRequest("PayrollIssueMonth").Item & oRequest("PayrollIssueDay").Item & oRequest("PayrollIssueYear").Item 'Fecha de aplicación
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						lCounter = 2
						Do While Not oRecordset.EOF
							sRowContents = "2" 'Tipo de registro
							sRowContents = sRowContents & Right(("00000" & lCounter), Len("00000")) 'Número de secuencia
							sRowContents = sRowContents & Right(("000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("000000")) 'Número de empleado
							sRowContents = sRowContents & " "
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeLastName").Value) & "                              "), Len("                              ")) 'Apellido paterno
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
							Err.Clear
							sRowContents = sRowContents & Left((sTemp & "                    "), Len("                    ")) 'Apellido materno
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeName").Value) & "                              "), Len("                              ")) 'Nombre
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("AccountNumber").Value) & "                "), Len("                ")) 'Número de cuenta
							sRowContents = sRowContents & Right(("000000000000000000" & Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)), Len("000000000000000000")) 'Importe
							dTotal = dTotal + Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100)
							lCounter = lCounter + 1
							lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						sRowContents = "3" 'Tipo de registro
						sRowContents = sRowContents & Right(("00000" & lCounter), Len("00000")) 'Número de secuencia
						sRowContents = sRowContents & Right(("00000" & lCounter-2), Len("00000")) 'Total de registros
							sRowContents = sRowContents & Right(("000000000000000000" & dTotal), Len("000000000000000000")) 'Importe total
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					End If
				Case 24 'HSBC
					Do While Not oRecordset.EOF
						sRowContents = Right(("0000000000" & CStr(oRecordset.Fields("AccountNumber").Value)), Len("0000000000")) 'Cuenta
						sRowContents = sRowContents & Right(("00000000000000000" & (Int(CDbl(oRecordset.Fields("CheckAmount").Value) * 100) / 100)), Len("00000000000000000")) 'Importe
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
				Case 25 'SPEI
					asRecordSpei = oRecordset.GetRows()
					dTotal = 0
					For iIndex = 0 To UBound(asRecordSpei,2)
						dTotal = dTotal + CDbl(asRecordSpei(0,iIndex))
					Next
					'Regsitro de control
					sRowContents = "1"
					sRowContents = sRowContents & Right(("000000000000" & CStr(oRequest("Field02").Item)), Len("000000000000"))
					sRowContents = sRowContents & oRequest("PayrollDepositDay").Item & oRequest("PayrollDepositMonth").Item & Right(oRequest("PayrollDepositYear").Item, Len("00")) 'Fecha de pago
                    'sRowContents = sRowContents & Right("00" & oRequest("PayrollIssueDay").Item,2) & Right("00" & oRequest("PayrollIssueMonth").Item,2) & Mid(oRequest("PayrollIssueYear").Item,3,2)
					sRowContents = sRowContents & Right("0000" & oRequest("FileNumber").Item,4)
					sRowContents = sRowContents & "ISSSTE                              "
					sRowContents = sRowContents & Left(sPayrollName & "                    ",20)
					sRowContents = sRowContents & "07"
					sRowContents = sRowContents & "                                        "
					sRowContents = sRowContents & "C"
					sRowContents = sRowContents & "0"
					sRowContents = sRowContents & "0"
					lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					'Registro global
					sRowContents = "2"
					sRowContents = sRowContents & "1"
					sRowContents = sRowContents & "001"
					sRowContents = sRowContents & Right("000000000000000000" & CStr(dTotal*100),18)
					sRowContents = sRowContents & "01"
					sRowContents = sRowContents & Right("0000" & oRequest("Field03").Item,4)
					sRowContents = sRowContents & Right("00000000000000000000" & oRequest("Field04").Item,20)
					sRowContents = sRowContents & "                    "
					lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					'Registro de cargos o abonos individuales
					For iIndex = 0 To UBound(asRecordSpei,2)
						sRowContents = "3"
						sRowContents = sRowContents & "0"
						sRowContents = sRowContents & "001"
						sRowContents = sRowContents & Right("000000000000000000" & CStr(CDbl(asRecordSpei(0,iIndex))*100),18)
						sRowContents = sRowContents & "01"
						sRowContents = sRowContents & Right("00000000000000000000" & CStr(asRecordSpei(2,iIndex)),20)
						sRowContents = sRowContents & "PAGO NOMINA                             "
						sTemp = asRecordSpei(4,iIndex) & "," & asRecordSpei(5,iIndex)
						If Len(asRecordSpei(6,iIndex)) <> 0 Then sTemp = sTemp & "/" & asRecordSpei(6,iIndex)
						sRowContents = sRowContents & Left(sTemp & "                                                  ",55)
						sRowContents = sRowContents & "                                        "
						sRowContents = sRowContents & "                        "
						sRowContents = sRowContents & Left( "0"&CStr(asRecordSpei(2,iIndex)),4)
                        'sRowContents = sRowContents & Right("0000" & CStr(asRecordSpei(1,iIndex)),4)
                        sRowContents = sRowContents & "0"& oRequest("PayrollDepositDay").Item & oRequest("PayrollDepositMonth").Item & Right(oRequest("PayrollDepositYear").Item, Len("00")) 'Fecha de pago
                    	'sRowContents = sRowContents & Right("0000000" & oRequest("SpeiRef").Item,7)
						sRowContents = sRowContents & "00"
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						oRecordset.MoveNext
					Next
					'Registro de totales
					sRowContents = "4"
					sRowContents = sRowContents & "001"
					sRowContents = sRowContents & Right("000000000000000000" & CStr(UBound(asRecordSpei,2) + 1),6)
					sRowContents = sRowContents & Right("000000000000000000" & CStr(dTotal*100),18)
					sRowContents = sRowContents & "000001"
					sRowContents = sRowContents & Right("000000000000000000" & CStr(dTotal*100),18)
					lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
			End Select

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".txt", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	Else
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1474 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1475(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for bank deposits
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1475"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sPayrollDate
	Dim bPayrollIsClosed
	Dim oRecordset
	Dim sRowContents
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim lErrorNumber

	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	sFilePath = Server.MapPath(sFilePath)
	Response.Flush()

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(sCondition, "Areas.", "Areas2."), "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList.")
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	sErrorDescription = "No se pudieron obtener los depósitos bancarios."
	If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		sCondition = Replace(sCondition, "PaymentCenters.", "BeneficiaryNumber.")
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End EmployeeLastName2, Areas2.AreaShortName From Payments, EmployeesBeneficiariesLKP, Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (-2,-1,1)) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."),"BankAccounts.","EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, Areas2.AreaShortName From Payments, EmployeesBeneficiariesLKP, Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones, BankAccounts Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (-2,-1,1)) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, Case When BeneficiaryLastName2 Is Null Then ' ' Else BeneficiaryLastName2 End EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, EmployeesBeneficiariesLKP, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones, BankAccounts Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, EmployeesBeneficiariesLKP, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones, BankAccounts Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
		End If
	ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		sCondition = Replace(sCondition, "PaymentCenters.", "CreditorNumber.")
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesCreditorsLKP.CreditorNumber EmployeeNumber, CreditorName EmployeeName, CreditorLastName EmployeeLastName, CreditorLastName2 EmployeeLastName2, Areas2.AreaShortName From Payments, EmployeesCreditorsLKP, Areas Areas1, Areas Areas2, Zones, Zones AreasZones2, Zones ParentZones Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Payments.StatusID In (-2,-1,1)) Order By EmployeesCreditorsLKP.CreditorNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesCreditorsLKP.CreditorNumber EmployeeNumber, CreditorName EmployeeName, CreditorLastName EmployeeLastName, CreditorLastName2 EmployeeLastName2, Areas2.AreaShortName From Payments, EmployeesCreditorsLKP, Areas Areas1, Areas Areas2, Zones, Zones AreasZones2, Zones ParentZones Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (Payments.StatusID In (-2,-1,1)) Order By EmployeesCreditorsLKP.CreditorNumber -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, EmployeesCreditorsLKP, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, EmployeesCreditorsLKP, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
		End If
	Else
		sCondition = Replace(sCondition, "PaymentCenters.", "EmployeesHistoryList.")
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas2.AreaShortName From Payments, Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (-2,-1,1)) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas2.AreaShortName From Payments, Employees, EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.PaymentDate=" & lPayrollID & ") And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (-2,-1,1)) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentDate, CheckNumber, CheckAmount, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentDate, CheckNumber, CheckAmount, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas2.AreaShortName From Payments, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Zones, Zones As AreasZones2, Zones As ParentZones Where (Payments.AccountID=BankAccounts.AccountID) And (Payments.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (PaymentDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.AccountNumber='.') And (BankAccounts.Active=1) And (Payments.StatusID In (-2,-1,1)) " & sCondition & " Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
		End If
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				sPayrollDate = Right(CStr(oRecordset.Fields("PaymentDate").Value), Len("DD")) & "/" & Mid(CStr(oRecordset.Fields("PaymentDate").Value), Len("YYYYM"), Len("MM")) & "/" & Left(CStr(oRecordset.Fields("PaymentDate").Value), Len("YYYY"))
				sRowContents = sPayrollDate
				sRowContents = sRowContents & "," & Right(("00000000" & CStr(oRecordset.Fields("CheckNumber").Value)), Len("00000000"))
				sRowContents = sRowContents & "," & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 4, True, False, False)
				sRowContents = sRowContents & "," & Right(("000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("000000"))
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)
				Else
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)
				End If
				sRowContents = sRowContents & "," & CStr(oRecordset.Fields("AreaShortName").Value)
				sRowContents = sRowContents & "," & sPayrollDate & "-nomina,01"

				lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".txt", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	Else
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1475 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1476_(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for the payment fees
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1476"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollStartDate
	Dim lPayrollDate_1
	Dim lCurrentID
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim sTemp
	Dim lLines_0
	Dim lLines_1
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim lErrorNumber

	sContents = GetFileContents(Server.MapPath("Templates\Report_1476.htm"), sErrorDescription)
	If Len(sContents) > 0 Then
		oStartDate = Now()
		sDate = GetSerialNumberForDate("")
		sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc"
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".doc", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		sFilePath = Server.MapPath(sFilePath)
		Response.Flush()

		Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
		lPayrollStartDate = CLng(Left(lForPayrollID, Len("000000")) & "01")
		lPayrollDate_1 = AddDaysToSerialDate(lForPayrollID, 1)

		sCondition = Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies", "EmployeesHistoryList"), "Payroll_YYYYMMDD", "Payroll_" & lPayrollID), "(Zones.", "(AreasZones.")
		If Len(oRequest("EmployeeID").Item) > 0 Then sCondition = sCondition & " And (Employees.EmployeeID In (" & Replace(oRequest("EmployeeID").Item, " ", "") & "))"
		If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'			sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
			sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		End If
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID>=600000) And (EmployeesHistoryList.EmployeeID<700000)"

		sErrorDescription = "No se pudieron obtener los pagos de honorarios."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Employees.EmployeeID=BankAccounts.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, IsDeduction, OrderInList", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Employees.EmployeeID=BankAccounts.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, IsDeduction, OrderInList -->" & vbNewLine

		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = AppendTextToFile(sFilePath, "<HTML>", sErrorDescription)
					lCurrentID = -2
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							If lCurrentID <> -2 Then
								If lLines_1 > lLines_0 Then
									For iIndex = lLines_1 To 4
										sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><DEDUCCIONES />")
									Next
								Else
									For iIndex = lLines_0 To 4
										sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><PERCEPCIONES />")
									Next
								End If
								sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "")
								sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "")
								lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
							End If
							sRowContents = sContents
							sRowContents = Replace(sRowContents, "<RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
							sRowContents = Replace(sRowContents, "<CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
							sRowContents = Replace(sRowContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)))
							Else
								sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)))
							End If
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("EmployeeAddress").Value)
							sTemp = sTemp & " C.P. " & CStr(oRecordset.Fields("EmployeeZipCode").Value)
							sRowContents = Replace(sRowContents, "<EMP_ADDRESS />", CleanStringForHTML(Replace(sTemp, vbNewLine, " ")))
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("EmployeePhone").Value)
							sRowContents = Replace(sRowContents, "<EMPLOYEE_PHONE />", CleanStringForHTML(sTemp))
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("EmployeeCity").Value)
							sRowContents = Replace(sRowContents, "<EMPLOYEE_CITY />", CleanStringForHTML(sTemp))
							sRowContents = Replace(sRowContents, "<EMPLOYEE_STATE />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
							sRowContents = Replace(sRowContents, "<PAYROLL_START_DATE />", UCase(DisplayShortDateFromSerialNumber(lPayrollStartDate, -1, -1, -1)))
							sRowContents = Replace(sRowContents, "<PAYROLL_DATE />", UCase(DisplayShortDateFromSerialNumber(lForPayrollID, -1, -1, -1)))
							sRowContents = Replace(sRowContents, "<PAYROLL_DATE_1 />", UCase(DisplayShortDateFromSerialNumber(lPayrollDate_1, -1, -1, -1)))
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							lLines_0 = 0
							lLines_1 = 0
						End If
						If CLng(oRecordset.Fields("ConceptID").Value) = 0 Then
							sRowContents = Replace(sRowContents, "<CONCEPT_0 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_0_AS_TEXT />", UCase(FormatNumberAsText(CDbl(oRecordset.Fields("TotalAmount").Value), True)))
						ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
							sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</FONT></TD></TR><PERCEPCIONES />")
							lLines_0 = lLines_0 + 1
						ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
							sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</FONT></TD></TR><DEDUCCIONES />")
							lLines_1 = lLines_1 + 1
						End If

						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					If lLines_1 > lLines_0 Then
						For iIndex = lLines_1 To 4
							sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><DEDUCCIONES />")
						Next
					Else
						For iIndex = lLines_0 To 4
							sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><PERCEPCIONES />")
						Next
					End If
					sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "")
					sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "")
					lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
				lErrorNumber = AppendTextToFile(sFilePath, "</HTML>", sErrorDescription)

				lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".doc", ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".doc", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		Else
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No se pudo abrir la plantilla del reporte."
	End If

	Set oRecordset = Nothing
	BuildReport1476 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1476(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for the payment fees
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1476"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollStartDate
	Dim lPayrollDate_1
	Dim lCurrentID
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim sTemp
	Dim lLines_0
	Dim lLines_1
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim lErrorNumber
	Dim lineCount
    Dim bPayrollIsClosed
    Dim tracePath

	sContents = GetFileContents(Server.MapPath("Templates\Report_1476.htm"), sErrorDescription)
	If Len(sContents) > 0 Then
		lineCount = 0
		oStartDate = Now()
		sDate = GetSerialNumberForDate("")
		sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".htm"
        tracePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
		Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".htm", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
		sFilePath = Server.MapPath(sFilePath)
        tracePath = Server.MapPath(tracePath)
		Response.Flush()

		Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
		lPayrollStartDate = CLng(Left(lForPayrollID, Len("000000")) & "01")
		lPayrollDate_1 = AddDaysToSerialDate(lForPayrollID, 1)
		
		If Len(oRequest("EmployeeID").Item) > 0 Then sCondition = sCondition & " And (Employees.EmployeeID In (" & Replace(oRequest("EmployeeID").Item, " ", "") & "))"
		'If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		'	sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		'	sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		'End If
		
		Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)
        If Not bPayrollIsClosed Then
            sCondition = Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies", "EmployeesHistoryList"), "Payroll_YYYYMMDD", "Payroll_" & lPayrollID), "(Zones.", "(AreasZones.")
        Else
            sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll."), "Companies.", "EmployeesHistoryListForPayroll."), "Payroll_YYYYMMDD", "Payroll_" & lPayrollID), "(Zones.", "(AreasZones.")
        End If

        sErrorDescription = "No se pudieron obtener los pagos de honorarios."
        If bPayrollIsClosed Then
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, EmployeesHistoryListForPayroll, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, IsDeduction, OrderInList", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
            'lErrorNumber = AppendTextToFile(tracePath, "Query: " & vbTab & "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, EmployeesHistoryListForPayroll, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, IsDeduction, OrderInList", sErrorDescription)
		    Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, EmployeesHistoryListForPayroll, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, IsDeduction, OrderInList -->" & vbNewLine
        Else
		    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Employees.EmployeeID=BankAccounts.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, IsDeduction, OrderInList", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		    Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesExtraInfo, States, BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Areas As ParentAreas, Areas As PaymentCenters, Zones As AreasZones, Zones As AreasZones2, Zones As ParentZones, Zones Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Employees.EmployeeID=EmployeesExtraInfo.EmployeeID) And (EmployeesExtraInfo.StateID=States.StateID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (Employees.EmployeeID=BankAccounts.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.PaymentCenterID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID) And (Areas.ZoneID=AreasZones.ZoneID) And (AreasZones.ParentID=AreasZones2.ZoneID) And (AreasZones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>=0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, EmployeeAddress, EmployeeCity, StateName, EmployeeZipCode, EmployeePhone, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Companies.CompanyShortName, ParentAreas.AreaCode, PaymentCenters.AreaCode, OrderInList Order By CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryList.EmployeeNumber, IsDeduction, OrderInList -->" & vbNewLine
        End If

		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = AppendTextToFile(sFilePath, "<HTML>", sErrorDescription)
				lCurrentID = -2
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID <> -2 Then
							If lLines_1 > lLines_0 Then
								For iIndex = lLines_1 To 4
									sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><DEDUCCIONES />")
								Next
							Else
								For iIndex = lLines_0 To 4
									sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><PERCEPCIONES />")
								Next
							End If
							sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "")
							sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "")
							lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
                            lineCount = lineCount + 1
                            'lErrorNumber = AppendTextToFile(tracePath, "Current: " & vbTab & lCurrentID  & vbTab & " LineCount: "  & vbTab & lineCount, sErrorDescription)
						End If
						sRowContents = sContents
						sRowContents = Replace(sRowContents, "<RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)))
						sRowContents = Replace(sRowContents, "<CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)))
						sRowContents = Replace(sRowContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)))
						Else
							sRowContents = Replace(sRowContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)))
						End If
						sTemp = ""
						sTemp = CStr(oRecordset.Fields("EmployeeAddress").Value)
						sTemp = sTemp & " C.P. " & CStr(oRecordset.Fields("EmployeeZipCode").Value)
						sRowContents = Replace(sRowContents, "<EMP_ADDRESS />", CleanStringForHTML(Replace(sTemp, vbNewLine, " ")))
						sTemp = ""
						sTemp = CStr(oRecordset.Fields("EmployeePhone").Value)
						sRowContents = Replace(sRowContents, "<EMPLOYEE_PHONE />", CleanStringForHTML(sTemp))
						sTemp = ""
						sTemp = CStr(oRecordset.Fields("EmployeeCity").Value)
						sRowContents = Replace(sRowContents, "<EMPLOYEE_CITY />", CleanStringForHTML(sTemp))
						sRowContents = Replace(sRowContents, "<EMPLOYEE_STATE />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
						sRowContents = Replace(sRowContents, "<PAYROLL_START_DATE />", UCase(DisplayShortDateFromSerialNumber(lPayrollStartDate, -1, -1, -1)))
						sRowContents = Replace(sRowContents, "<PAYROLL_DATE />", UCase(DisplayShortDateFromSerialNumber(lForPayrollID, -1, -1, -1)))
						sRowContents = Replace(sRowContents, "<PAYROLL_DATE_1 />", UCase(DisplayShortDateFromSerialNumber(lPayrollDate_1, -1, -1, -1)))
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						lLines_0 = 0
						lLines_1 = 0
					End If
					If CLng(oRecordset.Fields("ConceptID").Value) = 0 Then
						sRowContents = Replace(sRowContents, "<CONCEPT_0 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
						sRowContents = Replace(sRowContents, "<CONCEPT_0_AS_TEXT />", UCase(FormatNumberAsText(CDbl(oRecordset.Fields("TotalAmount").Value), True)))
					ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
						sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</FONT></TD></TR><PERCEPCIONES />")
						lLines_0 = lLines_0 + 1
					ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
						sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & "</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & "</FONT></TD></TR><DEDUCCIONES />")
						lLines_1 = lLines_1 + 1
					End If

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					If lLines_1 > lLines_0 Then
						For iIndex = lLines_1 To 2
							sRowContents = Replace(sRowContents, "<DEDUCCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><DEDUCCIONES />")
						Next
					Else
						For iIndex = lLines_0 To 2
							sRowContents = Replace(sRowContents, "<PERCEPCIONES />", "<TR><TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;</FONT></TD></TR><PERCEPCIONES />")
						Next
					End If
				Loop
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					lineCount = lineCount + 1
                    lErrorNumber = AppendTextToFile(tracePath, "Current: " & vbTab & lCurrentID  & vbTab & " LineCount: "  & vbTab & lineCount, sErrorDescription)
				End IF
				lErrorNumber = AppendTextToFile(sFilePath, "<hr /><br /> Total de recibos generados: " & lineCount & "<br /> </HTML>", sErrorDescription)
				lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".htm", ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".doc", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		Else
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No se pudo abrir la plantilla del reporte."
	End If

	Set oRecordset = Nothing
	BuildReport1476 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1477(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for the ISR amounts
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1477"
	Dim lForPayrollID
	Dim sCondition
	Dim sCondition2
	Dim sContents
	Dim oRecordset
	Dim asCLCs
	Dim adTotals
	Dim dTotal
	Dim sCurrentID
	Dim iIndex
	Dim jIndex
	Dim sNames
	Dim asParameters
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lReportID
	Dim lErrorNumber
    Dim lPeriodID
    Dim asPeriods

    asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
    lPeriodID = oRequest("YearID").Item & "0" & asPeriods(CInt(oRequest("MonthID").Item))

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "EmployeesHistoryList."),"EmployeesHistoryList.","EmployeesHistoryListForPayroll.")
	lForPayrollID = oRequest("YearID").Item & Right(("0" & oRequest("MonthID").Item), Len("00"))

	sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
                                                     
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CLC.PayrollID, PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, ConceptID, Sum(ConceptAmount) As TotalAmount From PayrollsCLCs As CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (ConceptID In (48,55,110)) " & sCondition & " Group By CLC.PayrollID, PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, ConceptID Order By CLC.PayrollID, CLC.PayrollCLC", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select CLC.PayrollID, PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, ConceptID, Sum(ConceptAmount) TotalAmount From PayrollsCLCs CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones Zones2, Zones ParentZones Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (ConceptID In (48,55,110)) " & sCondition & " Group By CLC.PayrollID, PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, ConceptID Order By CLC.PayrollID, CLC.PayrollCLC -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".xls", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			sFilePath = Server.MapPath(sFilePath)
			Response.Flush()

			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1477.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
			sContents = Replace(sContents, "<FILTER_MONTH />", UCase(CleanStringForHTML(asMonthNames_es(CInt(oRequest("MonthID").Item)))))
			sContents = Replace(sContents, "<FILTER_YEAR />", oRequest("YearID").Item)
			sNames = ""
			If Len(oRequest("CompanyID").Item) > 0 Then Call GetNameFromTable(oADODBConnection, "Companies", CLng(oRequest("CompanyID").Item), "", ", ", sNames, "")
			sContents = Replace(sContents, "<COMPANY_NAME />", UCase(CleanStringForHTML(sNames)))
			lErrorNumber = AppendTextToFile(sFilePath, sContents, sErrorDescription)

			sCurrentID = ""
			asCLCs = ""
			adTotals = Split(",", ",")
			adTotals(0) = Split("0,0,0,0", ",")
			adTotals(1) = Split("0,0,0,0", ",")
			For iIndex = 0 To UBound(adTotals(0))
				adTotals(0)(iIndex) = 0
				adTotals(1)(iIndex) = 0
			Next
			Do While Not oRecordset.EOF
				If StrComp(sCurrentID, CStr(oRecordset.Fields("PayrollCLC").Value), vbBinaryCompare) <> 0 Then
					If Len(sCurrentID) > 0 Then
						asCLCs = Replace(asCLCs, "<CONCEPT_110 />", "0.00")
						asCLCs = Replace(asCLCs, "<CONCEPT_48 />", "0.00")
						asCLCs = Replace(asCLCs, "<CONCEPT_55 />", "0.00")
						asCLCs = Replace(asCLCs, "<CONCEPT_00 />", "0.00")
						asCLCs = asCLCs & LIST_SEPARATOR
					End If
					'If CLng(oRecordset.Fields("PayrollTypeID").Value) <> 1 Then
						asCLCs = asCLCs & CleanStringForHTML(CStr(oRecordset.Fields("PayrollName").Value))
					'Else
					'	asCLCs = asCLCs & "Ordinaria"
					'End If
					asCLCs = asCLCs & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1)
					asCLCs = asCLCs & TABLE_SEPARATOR & CStr(oRecordset.Fields("FilterParameters").Value)
					asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_110 />"
					asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_48 />"
					asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_55 />"
					asCLCs = asCLCs & TABLE_SEPARATOR & "<CONCEPT_00 />"
					asCLCs = asCLCs & TABLE_SEPARATOR & "&nbsp; "
					sCurrentID = CStr(oRecordset.Fields("PayrollCLC").Value)
				End If
				asCLCs = Replace(asCLCs, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
				Select Case CLng(oRecordset.Fields("ConceptID").Value)
					Case 110
						adTotals(0)(0) = adTotals(0)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Case 48
						adTotals(0)(1) = adTotals(0)(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Case 55
						adTotals(0)(2) = adTotals(0)(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Case 00
						adTotals(0)(3) = adTotals(0)(3) + CDbl(oRecordset.Fields("TotalAmount").Value)
				End Select

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			asCLCs = Replace(asCLCs, "<CONCEPT_110 />", "0.00")
			asCLCs = Replace(asCLCs, "<CONCEPT_48 />", "0.00")
			asCLCs = Replace(asCLCs, "<CONCEPT_55 />", "0.00")
			asCLCs = Replace(asCLCs, "<CONCEPT_00 />", "0.00")
			asCLCs = Split(asCLCs, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asCLCs)
				asCLCs(iIndex) = Split(asCLCs(iIndex), TABLE_SEPARATOR)
				sNames = ""
				sNames = "<COMPANIES /><EMPLOYEE_TYPES /><BANKS /><CHECK_CONCEPTS /><AREAS /><PAYMENT_CENTERS /><ZONES /><EMPLOYEES />"
				asParameters = Split(asCLCs(iIndex)(2), " And ")
				For jIndex = 0 To UBound(asParameters)
					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.EmployeeID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID In (", ""), "))", "")
						sNames = Replace(sNames, "<EMPLOYEES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.CompanyID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.CompanyID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Companies", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<COMPANIES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.EmployeeTypeID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeTypeID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "EmployeeTypes", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<EMPLOYEE_TYPES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "Areas.AreaID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<AREAS />", sCondition2 & "&nbsp;")
					End If
					If InStr(1, asParameters(jIndex), "Areas.AreaPath ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaPath Like ""%,", ""), ",%"")", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<AREAS />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.PaymentCenterID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.PaymentCenterID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<PAYMENT_CENTERS />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "ParentZones.ZoneID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(ParentZones.ZoneID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<ZONES />", sCondition2 & "&nbsp;")
					End If
					If InStr(1, asParameters(jIndex), "Zones.ZonePath ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Zones.ZonePath Like ""%,", ""), ",%"")", "")
						Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<ZONES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "BankAccounts.BankID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(BankAccounts.BankID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Banks", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<BANKS />", sCondition2 & "&nbsp;")
					End If

					If StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0 Then
						If UBound(asParameters) <= jIndex Then
							sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con cheque y depósitos&nbsp;")
						Else
							If (StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber<>""."")", vbBinaryCompare) = 0) Then
								sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con depósito&nbsp;")
							ElseIf (StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber=""."")", vbBinaryCompare) = 0) Then
								sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con cheque&nbsp;")
							End If
						End If
					ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID>=600000)", vbBinaryCompare) = 0 Then
						sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Honorarios&nbsp;")
					ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID>=700000)", vbBinaryCompare) = 0 Then
						sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Pensión alimenticia&nbsp;")
					End If
				Next
				asCLCs(iIndex)(2) = asCLCs(iIndex)(0) & " " & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sNames, "<EMPLOYEES />", ""), "<COMPANIES />", ""), "<EMPLOYEE_TYPES />", ""), "<AREAS />", ""), "<PAYMENT_CENTERS />", ""), "<ZONES />", ""), "<BANKS />", ""), "<CHECK_CONCEPTS />", "")
				asCLCs(iIndex)(0) = "&nbsp;"
				
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asCLCs(iIndex), True, ""), sErrorDescription)
			Next

			sContents = GetFileContents(Server.MapPath("Templates\FooterForReport_1477.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SUBTOTAL_110 />", FormatNumber(adTotals(0)(0), 2, True, False, True))
			sContents = Replace(sContents, "<SUBTOTAL_48 />", FormatNumber(adTotals(0)(1), 2, True, False, True))
			sContents = Replace(sContents, "<SUBTOTAL_55 />", FormatNumber(adTotals(0)(2), 2, True, False, True))
			sContents = Replace(sContents, "<SUBTOTAL_00 />", FormatNumber(adTotals(0)(3), 2, True, False, True))

			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
                                                             
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, -Sum(ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (ConceptID In (48,55,110)) " & sCondition & " Group By ConceptID Order By Payroll_" & oRequest("YearID").Item & ".ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptID, -Sum(ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (ConceptID In (48,55,110)) " & sCondition & " Group By ConceptID Order By Payroll_" & oRequest("YearID").Item & ".ConceptID -->" & vbNewLine
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 110
							adTotals(1)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 48
							adTotals(1)(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 55
							adTotals(1)(2) = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 00
							adTotals(1)(3) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
			End If
			sContents = Replace(sContents, "<TOTAL_CANCEL_110 />", FormatNumber(adTotals(1)(0), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_CANCEL_48 />", FormatNumber(adTotals(1)(1), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_CANCEL_55 />", FormatNumber(adTotals(1)(2), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_CANCEL_00 />", FormatNumber(adTotals(1)(3), 2, True, False, True))

			sContents = Replace(sContents, "<TOTAL_110 />", FormatNumber(adTotals(0)(0) - adTotals(1)(0), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_48 />", FormatNumber(adTotals(0)(1) - adTotals(1)(1), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_55 />", FormatNumber(adTotals(0)(2) - adTotals(1)(2), 2, True, False, True))
			sContents = Replace(sContents, "<TOTAL_00 />", FormatNumber(adTotals(0)(3) - adTotals(1)(3), 2, True, False, True))

			dTotal = 0
			adTotals(0)(1) = -adTotals(0)(1)
			adTotals(1)(1) = -adTotals(1)(1)
			For iIndex = 0 To UBound(adTotals(0))
				dTotal = dTotal + (adTotals(0)(iIndex) - adTotals(1)(iIndex))
			Next
			sContents = Replace(sContents, "<TOTAL />", FormatNumber(dTotal, 2, True, False, True))
			lErrorNumber = AppendTextToFile(sFilePath, sContents, sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".xls", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".xls", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1477 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1478(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To build the file for the ISN amounts
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1478"
	Dim lForPayrollID
	Dim dTaxAmount
	Dim sCondition
	Dim sCondition2
	Dim sContents
	Dim oRecordset
	Dim asCLCs
	Dim asConcepts
	Dim asCancelConcepts
	Dim iMax
	Dim adTotals
	Dim dTotal
	Dim iIndex
	Dim jIndex
	Dim sNames
	Dim asParameters
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lReportID
	Dim lErrorNumber
    Dim lPeriodID
    Dim asPeriods

    asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
    lPeriodID = oRequest("YearID").Item & "0" & asPeriods(CInt(oRequest("MonthID").Item))

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "EmployeesHistoryList."),"EmployeesHistoryList.","EmployeesHistoryListForPayroll.")
	lForPayrollID = oRequest("YearID").Item & Right(("0" & oRequest("MonthID").Item), Len("00"))
	dTaxAmount = CDbl(oRequest("TaxAmount").Item)

	sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
	                                                 
    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From PayrollsCLCs As CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 0) " & sCondition & " Group By Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, CLC.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By Payrolls.PayrollID, CLC.PayrollCLC", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) TotalAmount From PayrollsCLCs CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones Zones2, Zones ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 0) " & sCondition & " Group By Payrolls.PayrollID, Payrolls.PayrollName, Payrolls.PayrollTypeID, CLC.PayrollCLC, CLC.FilterParameters, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By Payrolls.PayrollID, CLC.PayrollCLC -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".xls", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			sFilePath = Server.MapPath(sFilePath)
			Response.Flush()

			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1478.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
			sContents = Replace(sContents, "<FILTER_MONTH />", UCase(CleanStringForHTML(asMonthNames_es(CInt(oRequest("MonthID").Item)))))
			sContents = Replace(sContents, "<FILTER_YEAR />", oRequest("YearID").Item)
			sNames = ""
			If Len(oRequest("CompanyID").Item) > 0 Then Call GetNameFromTable(oADODBConnection, "Companies", CLng(oRequest("CompanyID").Item), "", ", ", sNames, "")
			sContents = Replace(sContents, "<COMPANY_NAME />", UCase(CleanStringForHTML(sNames)))
			sContents = Replace(sContents, "<TAX_AMOUNT />", dTaxAmount)
			lErrorNumber = AppendTextToFile(sFilePath, sContents, sErrorDescription)

			asCLCs = ""
			adTotals = Split(",", ",")
			adTotals(0) = Split("0,0,0,0", ",")
			adTotals(1) = Split("0,0,0,0", ",")
			For iIndex = 0 To UBound(adTotals(0))
				adTotals(0)(iIndex) = 0
				adTotals(1)(iIndex) = 0
			Next
			Do While Not oRecordset.EOF
				'If CLng(oRecordset.Fields("PayrollTypeID").Value) <> 1 Then
					asCLCs = asCLCs & CleanStringForHTML(CStr(oRecordset.Fields("PayrollName").Value))
				'Else
				'	asCLCs = asCLCs & "Ordinaria"
				'End If
				asCLCs = asCLCs & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1)
				asCLCs = asCLCs & TABLE_SEPARATOR & CStr(oRecordset.Fields("FilterParameters").Value)
				asCLCs = asCLCs & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
				asCLCs = asCLCs & LIST_SEPARATOR
				adTotals(0)(0) = adTotals(0)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close

			asCLCs = Split(asCLCs, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asCLCs) - 1
				asCLCs(iIndex) = Split(asCLCs(iIndex), TABLE_SEPARATOR)
				sNames = ""
				sNames = "<COMPANIES /><EMPLOYEE_TYPES /><BANKS /><CHECK_CONCEPTS /><AREAS /><PAYMENT_CENTERS /><ZONES /><EMPLOYEES />"
				asParameters = Split(asCLCs(iIndex)(2), " And ")
				For jIndex = 0 To UBound(asParameters)
					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.EmployeeID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID In (", ""), "))", "")
						sNames = Replace(sNames, "<EMPLOYEES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.CompanyID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.CompanyID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Companies", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<COMPANIES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.EmployeeTypeID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeTypeID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "EmployeeTypes", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<EMPLOYEE_TYPES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "Areas.AreaID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<AREAS />", sCondition2 & "&nbsp;")
					End If
					If InStr(1, asParameters(jIndex), "Areas.AreaPath ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Areas.AreaPath Like ""%,", ""), ",%"")", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<AREAS />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "EmployeesHistoryListForPayroll.PaymentCenterID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(EmployeesHistoryListForPayroll.PaymentCenterID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Areas", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<PAYMENT_CENTERS />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "ParentZones.ZoneID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(ParentZones.ZoneID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<ZONES />", sCondition2 & "&nbsp;")
					End If
					If InStr(1, asParameters(jIndex), "Zones.ZonePath ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(Zones.ZonePath Like ""%,", ""), ",%"")", "")
						Call GetNameFromTable(oADODBConnection, "Zones", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<ZONES />", sCondition2 & "&nbsp;")
					End If

					If InStr(1, asParameters(jIndex), "BankAccounts.BankID ", vbBinaryCompare) > 0 Then
						sCondition2 = Replace(Replace(asParameters(jIndex), "(BankAccounts.BankID In (", ""), "))", "")
						Call GetNameFromTable(oADODBConnection, "Banks", sCondition2, "", ",<BR />", sCondition2, sErrorDescription)
						sNames = Replace(sNames, "<BANKS />", sCondition2 & "&nbsp;")
					End If

					If StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0 Then
						If UBound(asParameters) <= jIndex Then
							sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con cheque y depósitos&nbsp;")
						Else
							If (StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber<>""."")", vbBinaryCompare) = 0) Then
								sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con depósito&nbsp;")
							ElseIf (StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID<600000)", vbBinaryCompare) = 0) And (StrComp(asParameters(jIndex + 1), "(BankAccounts.AccountNumber=""."")", vbBinaryCompare) = 0) Then
								sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Empleados con cheque&nbsp;")
							End If
						End If
					ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID>=600000)", vbBinaryCompare) = 0 Then
						sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Honorarios&nbsp;")
					ElseIf StrComp(asParameters(jIndex), "(EmployeesHistoryListForPayroll.EmployeeID>=700000)", vbBinaryCompare) = 0 Then
						sNames = Replace(sNames, "<CHECK_CONCEPTS />", "Pensión alimenticia&nbsp;")
					End If
				Next
				asCLCs(iIndex)(2) = asCLCs(iIndex)(0) & " " & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sNames, "<EMPLOYEES />", ""), "<COMPANIES />", ""), "<EMPLOYEE_TYPES />", ""), "<AREAS />", ""), "<PAYMENT_CENTERS />", ""), "<ZONES />", ""), "<BANKS />", ""), "<CHECK_CONCEPTS />", "")
				asCLCs(iIndex)(0) = "&nbsp;"
			Next

			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
                                                             
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From PayrollsCLCs As CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID)  And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 77) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By OrderInList, ConceptShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) TotalAmount From PayrollsCLCs CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas PaymentCenters, Zones, Zones Zones2, Zones ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID)  And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 77) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By OrderInList, ConceptShortName -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					adTotals(0)(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
				End If
				oRecordset.Close
			End If
			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From PayrollsCLCs As CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount<=0)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount>0))) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By OrderInList, ConceptShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) TotalAmount From PayrollsCLCs CLC, Payrolls, Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Areas PaymentCenters, Zones, Zones Zones2, Zones ParentZones, Concepts Where (PayrollCode = '" & lPeriodID & "') And (Payrolls.PayrollID = CLC.PayrollID) And (EmployeesHistoryListforPayroll.PayrollID = CLC.PayrollID) And (EmployeesHistoryListForPayroll.EmployeeID = CLC.EmployeeID) And (EmployeesHistoryListforPayroll.AreaID = Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID = PaymentCenters.AreaID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = Concepts.ConceptID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Payrolls.PayrollTypeID <> 0) And (Payroll_" & oRequest("YearID").Item & ".RecordDate = CLC.PayrollID) And (Payroll_" & oRequest("YearID").Item & ".EmployeeID = CLC.EmployeeID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount<=0)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount>0))) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID Order By OrderInList, ConceptShortName -->" & vbNewLine
			If lErrorNumber = 0 Then
				asConcepts = ""
				Do While Not oRecordset.EOF
					asConcepts = asConcepts & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & " " & CStr(oRecordset.Fields("ConceptName").Value))
					asConcepts = asConcepts & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					asConcepts = asConcepts & LIST_SEPARATOR
					adTotals(0)(3) = adTotals(0)(3) + CDbl(oRecordset.Fields("TotalAmount").Value)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			asConcepts = Split(asConcepts, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asConcepts)
				asConcepts(iIndex) = Split(asConcepts(iIndex), TABLE_SEPARATOR)
			Next
            'Desde aquí
			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 0) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 0) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					adTotals(1)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
				End If
				oRecordset.Close
			End If

			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 77) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (Payroll_" & oRequest("YearID").Item & ".ConceptID = 77) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					adTotals(1)(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
				End If
				oRecordset.Close
			End If

			sErrorDescription = "No se pudieron obtener los depósitos bloqueados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount<=0)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount>0))) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID, -Sum(Payroll_" & oRequest("YearID").Item & ".ConceptAmount) As TotalAmount From Payroll_" & oRequest("YearID").Item & ", EmployeesHistoryListForPayroll, Areas, Zones, Zones As Zones2, Zones As ParentZones, Concepts Where (Payroll_" & oRequest("YearID").Item & ".RecordDate >= " & lForPayrollID & "000) And (Payroll_" & oRequest("YearID").Item & ".RecordDate <= " & lForPayrollID & "990) And (Payroll_" & oRequest("YearID").Item & ".RecordID>=" & lForPayrollID & "00) And (Payroll_" & oRequest("YearID").Item & ".RecordID<=" & lForPayrollID & "99) And (Payroll_" & oRequest("YearID").Item & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryListForPayroll.PAYROLLID = Payroll_" & oRequest("YearID").Item & ".RecordDate) And (EmployeesHistoryListForPayroll.EmployeeID = Payroll_" & oRequest("YearID").Item & ".EmployeeID) And (EmployeesHistoryListForPayroll.AreaID = Areas.AreaID) And (Areas.ZoneID = Zones.ZoneID) And (Zones.ParentID = Zones2.ZoneID) And (Zones2.ParentID = ParentZones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount<=0)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount>0))) " & sCondition & " Group By ConceptShortName, ConceptName, OrderInList, Payroll_" & oRequest("YearID").Item & ".ConceptID  Order By Payroll_" & oRequest("YearID").Item & ".ConceptID -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					asCancelConcepts = asCancelConcepts & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & " " & CStr(oRecordset.Fields("ConceptName").Value))
					asCancelConcepts = asCancelConcepts & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					asCancelConcepts = asCancelConcepts & LIST_SEPARATOR
					adTotals(1)(3) = adTotals(1)(3) + CDbl(oRecordset.Fields("TotalAmount").Value)
				End If
				oRecordset.Close
			End If
			asCancelConcepts = Split(asCancelConcepts, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asCancelConcepts)
				asCancelConcepts(iIndex) = Split(asCancelConcepts(iIndex), TABLE_SEPARATOR)
			Next

			iMax = UBound(asCLCs) - 1
			If iMax < UBound(asConcepts) Then iMax = UBound(asConcepts) - 1
			If iMax < UBound(asCancelConcepts) Then iMax = UBound(asCancelConcepts) - 1
			For iIndex = 0 To iMax
				If iIndex < UBound(asCLCs) Then
					sRowContents = asCLCs(iIndex)(1)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(2)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(iIndex)(3)
				Else
					sRowContents = "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				If iIndex < UBound(asConcepts) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & asConcepts(iIndex)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & asConcepts(iIndex)(1)
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				If iIndex = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				If iIndex < UBound(asCancelConcepts) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & asCancelConcepts(iIndex)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & asCancelConcepts(iIndex)(1)
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
			Next

			adTotals(0)(2) = adTotals(0)(0) + adTotals(0)(1)
			adTotals(1)(2) = adTotals(1)(3) - adTotals(1)(1)

			sContents = GetFileContents(Server.MapPath("Templates\FooterForReport_1478.htm"), sErrorDescription)
			sContents = Replace(sContents, "<TAX_AMOUNT />", dTaxAmount)
			For iIndex = 0 To UBound(adTotals)
				For jIndex = 0 To UBound(adTotals(iIndex))
					sContents = Replace(sContents, "<TOTAL_" & iIndex & "_" & jIndex & " />", FormatNumber(adTotals(iIndex)(jIndex), 2, True, False, True))
				Next
			Next

			sContents = Replace(sContents, "<SUBTOTAL_0 />", FormatNumber((adTotals(0)(2) - adTotals(0)(3)), 2, True, False, True))
			sContents = Replace(sContents, "<SUBTOTAL_1 />", FormatNumber((adTotals(0)(2) - adTotals(0)(3) - adTotals(1)(0)), 2, True, False, True))
			sContents = Replace(sContents, "<SUBTOTAL_2 />", FormatNumber((adTotals(1)(0) - adTotals(1)(2)), 2, True, False, True))
			dTotal = (adTotals(0)(2) - adTotals(0)(3) - adTotals(1)(0)) * dTaxAmount / 100
			sContents = Replace(sContents, "<TOTAL />", FormatNumber(dTotal, 2, True, False, True))
			lErrorNumber = AppendTextToFile(sFilePath, sContents, sErrorDescription)

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".xls", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".xls", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1478 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1490(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll group by states and filtered
'         by banks
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1490"
	Dim sContents
	Dim sCondition
	Dim sDistinct
	Dim sField
	Dim sField2
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim oRecordset
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
	Const S_STATE_IDS = "1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,19,20,-1,21,22,23,24,25,26,27,28,29,30,31,32,9"
	Dim asStateIDs
	Dim sStateName

	If Len(oRequest("ZoneID").Item) > 0 Then
		asStateIDs = Replace(oRequest("ZoneID").Item, " ", "")
		asStateIDs = asStateIDs & ",1000"
		If (InStr(1, asStateIDs, ",9,", vbBinaryCompare) > 0) Then
			asStateIDs = Replace(asStateIDs, ",9,", ",") & ",9"
		ElseIf (InStr(1, asStateIDs, "9,", vbBinaryCompare) = 1) Then
			asStateIDs = Replace(asStateIDs, "9,", "", 1, 1, vbBinaryCompare) & ",9"
		ElseIf (InStr(1, asStateIDs, ",9", vbBinaryCompare) = (Len(asStateIDs) - 1)) Then
		Else
			asStateIDs = asStateIDs & ",1000"
		End If
		asStateIDs = Split(asStateIDs, ",")
	Else
		asStateIDs = Split(S_STATE_IDS, ",")
	End If
	adTotals = Split(",,,", ",")
	For iIndex = 0 To UBound(adTotals)
		adTotals(iIndex) = Split(",", ",")
		adTotals(iIndex)(0) = 0
		adTotals(iIndex)(1) = 0
	Next
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "(Areas.", "(Areas1."), "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "Concepts.", "Payroll_" & lPayrollID & "."), "EmployeeTypes.", "EmployeesHistoryList.")
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	If (InStr(1, sCondition, "Areas.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "Areas1.", vbBinaryCompare) > 0) Then
		'xxxxxxxxxxxxxxxxx
	End If

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	If (iConnectionType <> ACCESS) And (iConnectionType <> ACCESS_DSN) Then
		sDistinct = "Distinct "
	Else
		sDistinct = ""
	End If
	If Len(oRequest("ForWorkingCenter").Item) = 0 Then
		sField = "ZonesForPaymentCenter"
		sField2 = "PaymentCenters"
	Else
		sField = "Zones"
		sField2 = "Areas2"
	End If
	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1490.htm"), sErrorDescription)
	sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
	If Len(oRequest("ForWorkingCenter").Item) = 0 Then
		sContents = Replace(sContents, "<WORKING_CENTER />", "POR CENTRO DE PAGO")
	Else
		sContents = Replace(sContents, "<WORKING_CENTER />", "PRESUPUESTAL")
	End If
	sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
	sContents = Replace(sContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
	Response.Write sContents
	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Foráneo,Registros,Percepciones,Deducciones,Líquido", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "0", vbBinaryCompare) = 0) Then
			For iIndex = 0 To UBound(asStateIDs) - 1
				If CLng(asStateIDs(iIndex)) = -1 Then
					sStateName = "20A HOSP. REG. PDTE. JUAREZ OAXACA, OAX."
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						End If
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						End If
					Else
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
						Else
                            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID=38) " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
						End If
					End If
				Else
					Call GetNameFromTable(oADODBConnection, "States", asStateIDs(iIndex), "", "", sStateName, "")
					sErrorDescription = "No se pudieron obtener los montos pagados."
					If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						End If
					ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
						End If
					Else
						If bPayrollIsClosed Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
						Else
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
                            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & "," & asStateIDs(iIndex) & "," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
						End If
					End If
				End If
				If lErrorNumber = 0 Then
					For jIndex = 0 To UBound(adTotals)
						adTotals(jIndex)(0) = 0
					Next
					If Not oRecordset.EOF Then
						adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
						Do While Not oRecordset.EOF
							If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
								adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
								adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							Else
								adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
							End If
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
						oRecordset.Close
					End If
					sRowContents = CleanStringForHTML(sStateName)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
					adTotals(3)(1) = adTotals(3)(1) + adTotals(3)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
					adTotals(1)(1) = adTotals(1)(1) + adTotals(1)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
					adTotals(0)(1) = adTotals(0)(1) + adTotals(0)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
					adTotals(2)(1) = adTotals(2)(1) + adTotals(2)(0)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
			Next

			sRowContents = "<B>TOTAL FORÁNEOS</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(1), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(1), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(1), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(1), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If

		For iIndex = 0 To UBound(adTotals)
			adTotals(iIndex)(0) = 0
		Next
		If (Len(oRequest("StateType").Item) = 0) Or (StrComp(oRequest("StateType").Item, "1", vbBinaryCompare) = 0) Then
		If StrComp(asStateIDs(UBound(asStateIDs)), "9", vbBinaryCompare) = 0 Then
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				Else
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				End If
			End If
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						Do While Not oRecordset.EOF
							adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
							oRecordset.MoveNext
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
						Loop
					End If
					oRecordset.Close
				End If
			End If
			sRowContents = "LOCAL" & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "<B>TOTAL LOCAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				Else
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (AccountNumber='.') " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				End If
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "CHEQUE"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (EmployeesHistoryListForPayroll.AccountNumber<>'.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (EmployeesHistoryListForPayroll.AccountNumber<>'.') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				Else
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') And (BankAccounts.AccountNumber<>'.') " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				End If
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "DÉBITO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			For iIndex = 0 To UBound(adTotals)
				adTotals(iIndex)(0) = 0
			Next
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".RecordID=EmployeesBeneficiariesLKP.EmployeeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			ElseIf StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
				End If
			Else
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				Else
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Count(" & sDistinct & "Payroll_" & lPayrollID & ".EmployeeID) As TotalPayments, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ConceptID From Payroll_" & lPayrollID & ", BankAccounts, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (" & sField2 & ".ParentID<>38) And (" & sField & ".ZonePath Like '" & S_WILD_CHAR & ",9," & S_WILD_CHAR & "') " & sCondition & " Group By ConceptID Order By ConceptID Desc -->" & vbNewLine
				End If
			End If
			For jIndex = 0 To UBound(adTotals)
				adTotals(jIndex)(0) = 0
			Next
			If Not oRecordset.EOF Then
				adTotals(3)(0) = CLng(oRecordset.Fields("TotalPayments").Value)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("ConceptID").Value) > 0 Then
						adTotals(1)(0) = adTotals(1)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotals(2)(0) = adTotals(2)(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						adTotals(CLng(oRecordset.Fields("ConceptID").Value) + 2)(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
			sRowContents = "<B>TOTAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If
		End If

		If (Len(oRequest("StateType").Item) = 0) Then
			asRowContents = Split("&nbsp;,&nbsp;,&nbsp;,&nbsp;,&nbsp;", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "FORÁNEO"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(1), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(1), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(1), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(1), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "LOCAL"
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(3)(0), 0, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(1)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(0)(0), 2, True, False, True)
			sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(2)(0), 2, True, False, True)
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If

			sRowContents = "<B>TOTAL</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(3)(1) + adTotals(3)(0), 0, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(1)(1) + adTotals(1)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(0)(1) + adTotals(0)(0), 2, True, False, True) & "</B>"
			sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(2)(1) + adTotals(2)(0), 2, True, False, True) & "</B>"
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Else
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			End If
		End If
	Response.Write "</TABLE>"

	Set oRecordset = Nothing
	BuildReport1490 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1491(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll as text for the given concept
'         group by areas
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1491"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim lPayrollNumber
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sTemp
	Dim lCurrentID
	Dim sConceptShortName
	Dim alCounter
	Dim adTotal
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(sCondition, "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList.")
	oStartDate = Now()
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	Call GetNameFromTable(oADODBConnection, "ShortConcepts", oRequest("ConceptID").Item, "", "", sConceptShortName, sErrorDescription)
	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)
	If bPayrollIsClosed Then sCondition = Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.")

	alCounter = Split("0,0", ",")
	alCounter(0) = 0
	alCounter(1) = 0
	adTotal = Split("0,0", ",")
	adTotal(0) = 0
	adTotal(1) = 0
	Select Case CLng(oRequest("ConceptID").Item)
		Case 57
			sConceptShortName = "SH"
			sCondition = Replace(sCondition, "(Concepts.ConceptID In (" & oRequest("ConceptID").Item & "))", "(Concepts.ConceptID In (1," & oRequest("ConceptID").Item & "))")
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If bPayrollIsClosed Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
			End If
		Case 80
			lPayrollNumber = Right(("0000" & (CInt(GetPayrollNumber(lForPayrollID)) * 100) + CInt(Right(Left(lForPayrollID, Len("0000")), Len("00")))), Len("0000"))
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If bPayrollIsClosed Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaShortName, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
			End If
		Case 6364
			sConceptShortName = "63"
			sCondition = Replace(sCondition, "(Concepts.ConceptID In (" & oRequest("ConceptID").Item & "))", "(Concepts.ConceptID In (65,66))")
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If bPayrollIsClosed Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Case When ContractNumber Is Null Then ' ' Else ContractNumber End ContractNumber, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & " LEFT OUTER JOIN Credits On ((Payroll_" & lPayrollID & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Credits.CreditTypeID) And (Credits.StartDate<=" & lForPayrollID & ") And (Credits.EndDate>=" & lForPayrollID & ")), Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID, Credits.ContractNumber Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Case When ContractNumber Is Null Then ' ' Else ContractNumber End ContractNumber, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & " LEFT OUTER JOIN Credits On ((Payroll_" & lPayrollID & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Credits.CreditTypeID) And (Credits.StartDate<=" & lForPayrollID & ") And (Credits.EndDate>=" & lForPayrollID & ")), Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID, Credits.ContractNumber Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Case When ContractNumber Is Null Then ' ' Else ContractNumber End ContractNumber, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ",  LEFT OUTER JOIN Credits On ((Payroll_" & lPayrollID & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Credits.CreditTypeID) And (Credits.StartDate<=" & lForPayrollID & ") And (Credits.EndDate>=" & lForPayrollID & ")), Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID, Credits.ContractNumber Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Case When ContractNumber Is Null Then ' ' Else ContractNumber End ContractNumber, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ",  LEFT OUTER JOIN Credits On ((Payroll_" & lPayrollID & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Credits.CreditTypeID) And (Credits.StartDate<=" & lForPayrollID & ") And (Credits.EndDate>=" & lForPayrollID & ")), Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID, Credits.ContractNumber Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
			End If
		Case Else
			sCondition = Replace(sCondition, "(Concepts.ConceptID In (" & oRequest("ConceptID").Item & "))", "(Concepts.ConceptID In (1," & oRequest("ConceptID").Item & "))")
			sErrorDescription = "No se pudieron obtener los montos pagados."
			If bPayrollIsClosed Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By ConceptShortName, EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine
			End If
	End Select
	sConceptShortName = SizeText(sConceptShortName, " ", 2)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			If lErrorNumber = 0 Then
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

				lCurrentID = -2
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						If lCurrentID <> -2 Then
							If InStr(1, sRowContents, "<CONCEPT_?? />", vbBinaryCompare) = 0 Then
								Select Case CLng(oRequest("ConceptID").Item)
									Case 80
										'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000000000")
									Case 6364
										'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000000")
									Case 57
										sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "00000000")
										sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText(Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("00000000")), " ", 8, 1))
									Case Else
										sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "00000000")
										sRowContents = Replace(sRowContents, "<CONCEPT_86 />", "00000000")
										'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "00000000")
								End Select
								lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
							End If
						End If
						Select Case CLng(oRequest("ConceptID").Item)
							Case 80
								sRowContents = SizeText(CStr(oRecordset.Fields("EmployeeNumber").Value), " ", 6, 1) 'Empleado
								sTemp = CStr(oRecordset.Fields("EmployeeName").Value)
								sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
								Err.Clear
								sRowContents = sRowContents & SizeText(sTemp, " ", 30, 1) 'Nombre
								sRowContents = sRowContents & "<CONCEPT_?? />" 'Importe
								sRowContents = sRowContents & lPayrollNumber 'Quincena y año de pago
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("AreaShortName").Value), " ", 10, 1) 'Adscripción
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 10, 1) 'RFC
								sRowContents = sRowContents & sConceptShortName 'Concepto
							Case 6364
								sRowContents = "10007999" 'Tipo de registro, Clave1, Clave 2
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("ContractNumber").Value), " ", 6, 1) 'No. Póliza
								sRowContents = sRowContents & lPayrollNumber 'Fecha de pago
								sRowContents = sRowContents & Right(("000" & CStr(oRecordset.Fields("ConceptShortName").Value)), Len("000")) 'Concepto
								sRowContents = sRowContents & "P" 'Tipo de movimiento
								sRowContents = sRowContents & "<CONCEPT_?? />" 'Importe
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 13, 1) 'RFC
								sTemp = CStr(oRecordset.Fields("EmployeeName").Value)
								sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
								Err.Clear
								sRowContents = sRowContents & SizeText(sTemp, " ", 29, 1) 'Nombre
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeNumber").Value), " ", 6, 1) 'Empleado
								sRowContents = sRowContents & "                     " 'Espacios
							Case Else
								sRowContents = lPayrollNumber 'Fecha de pago
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 13, 1) 'RFC
								sTemp = CStr(oRecordset.Fields("EmployeeName").Value)
								sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
								Err.Clear
								sRowContents = sRowContents & SizeText(sTemp, " ", 35, 1) 'Nombre
								sRowContents = sRowContents & sConceptShortName 'Concepto
								sRowContents = sRowContents & "<CONCEPT_?? />" 'Importe
								sRowContents = sRowContents & "<CONCEPT_01 />" 'Sueldo base
								'sRowContents = sRowContents & "<CONCEPT_86 />" 'Seguro daños FOVISSSTE
								sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("EmployeeNumber").Value), " ", 6, 1) 'Empleado
						End Select
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 1
							If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
								sRowContents = Replace(sRowContents, "<CONCEPT_01 />", SizeText(Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("00000000")), " ", 8, 1))
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_01 />", SizeText("-" & Right(("0000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000")), " ", 8, 1))
							End If
						'Case 83
						'	If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
						'		sRowContents = Replace(sRowContents, "<CONCEPT_86 />", SizeText(Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("00000000")), " ", 8, 1))
						'		sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText(Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("00000000")), " ", 8, 1))
						'	Else
						'		sRowContents = Replace(sRowContents, "<CONCEPT_86 />", SizeText("-" & Right(("0000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000")), " ", 8, 1))
						'		sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText("-" & Right(("0000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000")), " ", 8, 1))
						'	End If
						'	alCounter(0) = alCounter(0) + 1
						'	adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 80
							If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText(Right(("0000000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000000")), " ", 10, 1))
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText("-" & Right(("000000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000000")), " ", 10, 1))
							End If
							alCounter(0) = alCounter(0) + 1
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 65, 66
							If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText(Right(("0000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000")), " ", 7, 1))
								alCounter(CLng(oRecordset.Fields("ConceptID").Value) - 65) = alCounter(CLng(oRecordset.Fields("ConceptID").Value) - 65) + 1
								adTotal(CLng(oRecordset.Fields("ConceptID").Value) - 65) = adTotal(CLng(oRecordset.Fields("ConceptID").Value) - 65) + CDbl(oRecordset.Fields("TotalAmount").Value)
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText("-" & Right(("000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("000000")), " ", 7, 1))
							End If
							lCurrentID = -3
						Case Else
							If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText(Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("00000000")), " ", 8, 1))
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_?? />", SizeText("-" & Right(("0000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value) * 100, 0, True, False, False)), Len("0000000")), " ", 8, 1))
							End If
							alCounter(0) = alCounter(0) + 1
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				If InStr(1, sRowContents, "<CONCEPT_?? />", vbBinaryCompare) = 0 Then
					Select Case CLng(oRequest("ConceptID").Item)
						Case 80
							'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000000000")
							lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
						Case 6364
							'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000000")
							lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
						Case Else
							sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "00000000")
							sRowContents = Replace(sRowContents, "<CONCEPT_86 />", "00000000")
							'sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "00000000")
							lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
					End Select
				End If
				oRecordset.Close
				Select Case CLng(oRequest("ConceptID").Item)
					Case 80
						sRowContents = "208429990" & sConceptShortName & "P" & Right(("00000000" & alCounter(0)), Len("00000000")) & Right(("000000000000" & FormatNumber((adTotal(0) * 100), 0, True, False, False)), Len("000000000000")) & SizeText("", " ", 68, 1) 'Total Concepto 80
						lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
					Case 6364
						sRowContents = "20842999063P" & Right(("00000000" & alCounter(0)), Len("00000000")) & Right(("000000000000" & FormatNumber((adTotal(0) * 100), 0, True, False, False)), Len("000000000000")) & SizeText("", " ", 68, 1) 'Total Concepto 63
						lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
						sRowContents = "20842999064P" & Right(("00000000" & alCounter(1)), Len("00000000")) & Right(("000000000000" & FormatNumber((adTotal(1) * 100), 0, True, False, False)), Len("000000000000")) & SizeText("", " ", 68, 1) 'Total Concepto 64
						lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
					Case Else
						sRowContents = "208429990" & sConceptShortName & "P" & Right(("00000000" & alCounter(0)), Len("00000000")) & Right(("000000000000" & FormatNumber((adTotal(0) * 100), 0, True, False, False)), Len("000000000000")) & SizeText("", " ", 68, 1) 'Total Concepto 80
						lErrorNumber = AppendTextToFile((sFilePath & "Terceros_" & sConceptShortName & ".txt"), sRowContents, sErrorDescription)
				End Select

				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1491 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1492(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the given concept group by areas
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1492"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim dTotal
	Dim dGranTotal
	Dim lCounter
	Dim lGranCounter
	Dim sCompanyName
	Dim sConceptName
	Dim sConceptShortName
	Dim sContents
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sCondition = Replace(Replace(Replace(sCondition, "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList."), "Companies.", "EmployeesHistoryList.")
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".doc")
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".doc", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)
	If bPayrollIsClosed Then sCondition = Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.")

	Call GetNameFromTable(oADODBConnection, "Companies", oRequest("CompanyID").Item, "", ", ", sCompanyName, "")
	Call GetNameFromTable(oADODBConnection, "ConceptsNames", oRequest("ConceptID").Item, "", "", sConceptName, "")
	Call GetNameFromTable(oADODBConnection, "ShortConcepts", oRequest("ConceptID").Item, "", "", sConceptShortName, "")
	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1492.htm"), sErrorDescription)
	sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
	sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))), -1, -1, -1))
	sContents = Replace(sContents, "<CURRENT_HOUR />", Hour(Time()) & ":" & Minute(Time()) & ":" & Second(Time()))
	sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1))
	sContents = Replace(sContents, "<CONCEPT_NAME />", CleanStringForHTML(sConceptName))
	sContents = Replace(sContents, "<CONCEPT_SHORT_NAME />", CleanStringForHTML(sConceptShortName))
	sContents = Replace(sContents, "<PAYROLL_NUMBER />", GetPayrollNumber(lForPayrollID))
	sContents = Replace(sContents, "<PAYROLL_YEAR />", Left(lForPayrollID, Len("0000")))
	lErrorNumber = AppendTextToFile(sFilePath, sContents, sErrorDescription)
	lErrorNumber = AppendTextToFile(sFilePath, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
		asColumnsTitles = Split("No. del empleado,Nombre,Adscripción,R.F.C.,Puesto,Nivel,Concepto (" & CleanStringForHTML(sConceptShortName) & ")", ",", -1, vbBinaryCompare)
		lErrorNumber = AppendTextToFile(sFilePath, GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)

		asCellAlignments = Split(",,,,,,RIGHT", ",", -1, vbBinaryCompare)
		dTotal = 0
		dGranTotal = 0
		lCounter = 0
		lGranCounter = 0
		sRowContents = "<SPAN COLS=""7"" /><B>EMPRESA: " & CleanStringForHTML(sCompanyName) & "</B>"
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)

		sErrorDescription = "No se pudieron obtener los montos pagados."
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID=1) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID=1) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID=1) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID=1) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName -->" & vbNewLine
		End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sRowContents = "<SPAN COLS=""7"" /><B>FUNCIONARIOS</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000")))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
					lCounter = lCounter + 1

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "<SPAN COLS=""5"" /><B>Total " & CleanStringForHTML(sCompanyName & ". FUNCIONARIOS") & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(lCounter, 0, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				dGranTotal = dGranTotal + dTotal
				lGranCounter = lGranCounter + lCounter
				dTotal = 0
				lCounter = 0
				asRowContents = Split("<SPAN COLS=""7"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
			End If
		End If

		sErrorDescription = "No se pudieron obtener los montos pagados."
		If bPayrollIsClosed Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID In (0,2,3,4,5,6)) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesHistoryListForPayroll, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID  & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID In (0,2,3,4,5,6)) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID In (0,2,3,4,5,6)) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, Case When EmployeeLastName2 Is Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Areas.AreaCode, PositionShortName, LevelShortName, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Companies, EmployeeTypes, Areas, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Positions, Levels Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.EmployeeTypeID In (0,2,3,4,5,6)) " & sCondition & " Group By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName Order By ConceptShortName, ConceptName, CompanyName, EmployeeTypes.EmployeeTypeID, EmployeeTypeName, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Areas.AreaCode, PositionShortName, LevelShortName -->" & vbNewLine
		End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sRowContents = "<SPAN COLS=""7"" /><B>OPERATIVOS</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("000")))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
					lCounter = lCounter + 1

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "<SPAN COLS=""5"" /><B>Total " & CleanStringForHTML(sCompanyName & ". OPERATIVOS") & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(lCounter, 0, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				dGranTotal = dGranTotal + dTotal
				lGranCounter = lGranCounter + lCounter
				dTotal = 0
				lCounter = 0
				asRowContents = Split("<SPAN COLS=""7"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
			End If
		End If

		sRowContents = "<SPAN COLS=""5"" /><B>Total " & CleanStringForHTML(sCompanyName) & "</B>"
		sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(lGranCounter, 0, True, False, True) & "</B>"
		sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dGranTotal, 2, True, False, True) & "</B>"
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = AppendTextToFile(sFilePath, GetTableRowText(asRowContents, True, ""), sErrorDescription)
	lErrorNumber = AppendTextToFile(sFilePath, "</TABLE>", sErrorDescription)

	lErrorNumber = ZipFolder(sFilePath, Replace(sFilePath, ".doc", ".zip"), sErrorDescription)
	If lErrorNumber = 0 Then
		Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
		sErrorDescription = "No se pudieron guardar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
	End If
	oEndDate = Now()
	If (lErrorNumber = 0) And B_USE_SMTP Then
		If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".doc", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
	End If

	Set oRecordset = Nothing
	BuildReport1492 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1494(oRequest, oADODBConnection, sConceptID, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de auditoria de movimientos
'Inputs:  oRequest, oADODBConnection, sConceptID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1494"
	Dim sHeaderContents
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim sConceptName
	Dim sTypeName
	Dim iIndex
	Dim asTotals
	Dim sCondition
	Dim sCurrentID
	Dim lCompanyID
	Dim sCompanyName
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	Select Case sConceptID
		Case "56,76,77"
			sConceptName = "SNTISSSTE y FSTSE"
			asColumnsTitles = "Concepto,Fecha,No. registros,Clave,Importe,Importe 25% FONAC,Total,Total F.S.T.S.E. 10%,Total S.N.T.I.S.S.S.T.E." & sConceptName
		Case "101"
			sConceptName = "SITISSSTE"
			asColumnsTitles = "Concepto,Fecha,No. registros,Clave,Importe,&nbsp;,Total,&nbsp;,Total " & sConceptName
	End Select
	asTotals = Split("0,0,0", ",")
	asTotals(0) = Split("0,0,0,0,0,0", ",")
	asTotals(1) = Split("0,0,0,0,0,0", ",")
	asTotals(2) = Split("0,0,0,0,0,0", ",")
	For iIndex = 0 To UBound(asTotals(0))
		asTotals(0)(iIndex) = 0
		asTotals(1)(iIndex) = 0
		asTotals(2)(iIndex) = 0
	Next

	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	If bPayrollIsClosed Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName, Count(Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayrollID, Len("0000"))& ", EmployeesHistoryListForPayroll, Concepts, Companies, EmployeeTypes Where (Payroll_" & Left(lPayrollID, Len("0000"))& ".ConceptID=Concepts.ConceptID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".RecordID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName Order By CompanyShortName, Companies.CompanyID, EmployeeTypes.TypeID Desc, Concepts.ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName, Count(Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayrollID, Len("0000"))& ", EmployeesHistoryListForPayroll, Concepts, Companies, EmployeeTypes Where (Payroll_" & Left(lPayrollID, Len("0000"))& ".ConceptID=Concepts.ConceptID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".RecordID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName Order By CompanyShortName, Companies.CompanyID, EmployeeTypes.TypeID Desc, Concepts.ConceptID -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName, Count(Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayrollID, Len("0000"))& ", EmployeesChangesLKP, EmployeesHistoryList, Concepts, Companies, EmployeeTypes Where (Payroll_" & Left(lPayrollID, Len("0000"))& ".ConceptID=Concepts.ConceptID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".RecordID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName Order By CompanyShortName, Companies.CompanyID, EmployeeTypes.TypeID Desc, Concepts.ConceptID", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName, Count(Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayrollID, Len("0000"))& ", EmployeesChangesLKP, EmployeesHistoryList, Concepts, Companies, EmployeeTypes Where (Payroll_" & Left(lPayrollID, Len("0000"))& ".ConceptID=Concepts.ConceptID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Payroll_" & Left(lPayrollID, Len("0000"))& ".RecordID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyShortName, CompanyName, EmployeeTypes.TypeID, Concepts.ConceptID, ConceptShortName Order By CompanyShortName, Companies.CompanyID, EmployeeTypes.TypeID Desc, Concepts.ConceptID -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1494.htm"), sErrorDescription)
			If Len(sHeaderContents) > 0 Then
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				sHeaderContents = Replace(sHeaderContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", GetPayrollNumber(lForPayrollID) & "/" & Left(lForPayrollID, Len("0000")))
				sHeaderContents = Replace(sHeaderContents, "<CONCEPT_NAME />", sConceptName)
				If bForExport Then sHeaderContents = Replace(sHeaderContents, "<BR />", vbNewLine)
			End If
			Response.Write sHeaderContents
			Select Case sConceptID
				Case "56,76,77"
					sConceptName = "54"
				Case "101"
					sConceptName = "CS"
			End Select
			Response.Write "<TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">"
				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,RIGHT,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				sCurrentID = ""
				lCompanyID = CLng(oRecordset.Fields("CompanyID").Value)
				sCompanyName = CStr(oRecordset.Fields("CompanyName").Value)
				Do While Not oRecordset.EOF
					If StrComp(sCurrentID, CStr(oRecordset.Fields("CompanyID").Value) & "," & CStr(oRecordset.Fields("TypeID").Value), vbBinaryCompare) <> 0 Then
						If Len(sCurrentID) > 0 Then
							sRowContents = CleanStringForHTML(sTypeName & " " & sCompanyName)
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(lForPayrollID)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(0), 0, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sConceptName)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(1), 2, True, False, True)
							If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
								asTotals(2)(2) = asTotals(2)(2) * 0.25
								asTotals(2)(3) = asTotals(2)(1) - asTotals(2)(2)
								asTotals(2)(4) = asTotals(2)(3) * 0.1
								asTotals(2)(5) = asTotals(2)(3) - asTotals(2)(4)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(2), 2, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(3), 2, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(4), 2, True, False, True)
								sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(5), 2, True, False, True)
							Else
								asTotals(2)(2) = 0
								asTotals(2)(3) = asTotals(2)(1)
								asTotals(2)(4) = 0
								asTotals(2)(5) = asTotals(2)(1)
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
							End If

							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							For iIndex = 0 To UBound(asTotals(0))
								asTotals(0)(iIndex) = asTotals(0)(iIndex) + asTotals(2)(iIndex)
								asTotals(1)(iIndex) = asTotals(1)(iIndex) + asTotals(2)(iIndex)
								asTotals(2)(iIndex) = 0
							Next
						End If
						sCurrentID = CStr(oRecordset.Fields("CompanyID").Value) & "," & CStr(oRecordset.Fields("TypeID").Value)
					End If
					If lCompanyID <> CLng(oRecordset.Fields("CompanyID").Value) Then
						sRowContents = "<B>" & CleanStringForHTML("Subtotal " & sCompanyName) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(0), 0 , True, False, True) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(1), 2 , True, False, True) & "</B>"
						If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(2), 2 , True, False, True) & "</B>"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(3), 2 , True, False, True) & "</B>"
						If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(4), 2 , True, False, True) & "</B>"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(5), 2 , True, False, True) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If

						For iIndex = 0 To UBound(asTotals(0))
							asTotals(0)(iIndex) = 0
						Next
						lCompanyID = CLng(oRecordset.Fields("CompanyID").Value)
						sCompanyName = CStr(oRecordset.Fields("CompanyName").Value)
					End If
					sTypeName = "Funcionarios"
					If CInt(oRecordset.Fields("TypeID").Value) = 1 Then sTypeName = "Operativos"
					If InStr(1, ",56,101,", "," & CStr(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0 Then
						asTotals(2)(0) = CLng(oRecordset.Fields("TotalCount").Value)
						asTotals(2)(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						asTotals(2)(2) = asTotals(2)(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sRowContents = CleanStringForHTML(sTypeName & " " & sCompanyName)
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(lForPayrollID)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(0), 0, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sConceptName)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(1), 2, True, False, True)
				asTotals(2)(2) = asTotals(2)(2) * 0.25
				If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
					asTotals(2)(3) = asTotals(2)(1) - asTotals(2)(2)
					asTotals(2)(4) = asTotals(2)(3) * 0.1
					asTotals(2)(5) = asTotals(2)(3) - asTotals(2)(4)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(2), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(3), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(4), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(asTotals(2)(5), 2, True, False, True)
				Else
					asTotals(2)(3) = asTotals(2)(1)
					asTotals(2)(4) = 0
					asTotals(2)(5) = asTotals(2)(1)
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				For iIndex = 0 To UBound(asTotals(0))
					asTotals(0)(iIndex) = asTotals(0)(iIndex) + asTotals(2)(iIndex)
					asTotals(1)(iIndex) = asTotals(1)(iIndex) + asTotals(2)(iIndex)
					asTotals(2)(iIndex) = 0
				Next

				sRowContents = "<B>" & CleanStringForHTML("Subtotal " & sCompanyName) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(0), 0 , True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(1), 2 , True, False, True) & "</B>"
				If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(2), 2 , True, False, True) & "</B>"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(3), 2 , True, False, True) & "</B>"
				If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(4), 2 , True, False, True) & "</B>"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(0)(5), 2 , True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "&nbsp;"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>TOTALES</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(0), 0 , True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(1), 2 , True, False, True) & "</B>"
				If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(2), 2 , True, False, True) & "</B>"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(3), 2 , True, False, True) & "</B>"
				If StrComp(sConceptID, "56,76,77", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(4), 2 , True, False, True) & "</B>"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(asTotals(1)(5), 2 , True, False, True) & "</B>"

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.Close
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1494 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1495(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de auditoria de movimientos
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1495"
	Dim sHeaderContents
	Dim oRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sCondition
	Dim sQuery
	Dim iNumEmp
	Dim iNumEmpAnt

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "XXX", "Registration")

	oStartDate = Now()
	sQuery = "Select Audit.EmployeeID, EmployeeName, EmployeeLastName, EmployeeLastName2, Audit.ConceptID, Audit.StartDate, AuditOperationShortName, AuditTypeName, AuditsDate, UserName, UserLastName" & _
			 " From Audit, AuditTypes, AuditOperationTypes, Employees, Users" & _
			 " Where Audit.EmployeeID = Employees.EmployeeID" & _
			 " And Audit.AuditTypeID = AuditTypes.AuditTypeID" & _
			 " And Audit.AuditOperationTypeID = AuditOperationTypes.AuditOperationTypeID" & _
			 " And Audit.AuditUserID = Users.UserID" & sCondition

	sErrorDescription = "No se pudieron obtener los registros de auditoria."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

				sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1494.htm"), sErrorDescription)
				If Len(sHeaderContents) > 0 Then
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sFileName & ".xls", sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido materno del empleado</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Concepto</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha del concepto</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Tipo de movimiento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha del movimiento</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del usuario</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Apellido del usuario</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sFileName & ".xls", sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					iNumEmp = CInt(oRecordset.Fields("EmployeeNumber").Value)
					sRowContents = "<TR>"
						If iNumEmp <> iNumEmpAnt Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & "" & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</FONT></TD>"
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">&nbps;</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("ConceptID").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("StartDate").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AuditOperationShortName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AuditTypeName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("AuditsDate").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UserName").Value) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CStr(oRecordset.Fields("UserLastName").Value) & "</FONT></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sFileName & ".xls", sRowContents, sErrorDescription)
					iNumEmpAnt = iNumEmp
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sFileName & ".xls", sRowContents, sErrorDescription)
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName & ".zip"), sErrorDescription)
				'If lErrorNumber = 0 Then
				'	Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				'	sErrorDescription = "No se pudo guardar la información del reporte."
				'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				'End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName & ".zip", CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1495 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1499(oRequest, oADODBConnection, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To get the paid amounts for the given employees and
'         the give payroll
'Inputs:  oRequest, oADODBConnection, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1499"
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim sContents
	Dim iStartPos
	Dim iEndPos
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	sErrorDescription = "No se pudieron obtener los montos de la nómina del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, PositionShortName, LevelShortName, Areas.EconomicZoneID, CheckNumber, ConceptShortName, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Areas, Positions, Levels Where (Payments.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".EmployeeID=" & lEmployeeID & ") Group By EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, PositionShortName, LevelShortName, Areas.EconomicZoneID, CheckNumber, ConceptShortName Order By ConceptShortName", "ReportsQueries1400cLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\Report_1499.htm"), sErrorDescription)
			If Len(sContents) > 0 Then
				sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sContents = Replace(sContents, "<EMPLOYEE_NAME />", CleanStringforHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value)))
				Else
					sContents = Replace(sContents, "<EMPLOYEE_NAME />", CleanStringforHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)))
				End If
				sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sContents = Replace(sContents, "<EMPLOYEE_RFC />", CleanStringforHTML(CStr(oRecordset.Fields("RFC").Value)))
				sContents = Replace(sContents, "<CHECK_NUMBER />", CleanStringforHTML(CStr(oRecordset.Fields("CheckNumber").Value)))
				sContents = Replace(sContents, "<AREA_CODE />", CleanStringforHTML(CStr(oRecordset.Fields("AreaCode").Value)))
				sContents = Replace(sContents, "<POSITION_SHORT_NAME />", CleanStringforHTML(CStr(oRecordset.Fields("PositionShortName").Value)))
				sContents = Replace(sContents, "<LEVEL_SHORT_NAME />", CleanStringforHTML(CStr(oRecordset.Fields("LevelShortName").Value)))
				sContents = Replace(sContents, "<ECONOMIC_ZONE />", CleanStringforHTML(CStr(oRecordset.Fields("EconomicZoneID").Value)))
				sContents = Replace(sContents, "<EMPLOYEE_NUMBER />", CleanStringforHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)))
				sContents = Replace(sContents, "<CURRENT_DATE />", DisplayNumericDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000"))))
				sContents = Replace(sContents, "<PAYROLL_START_DATE />", DisplayNumericDateFromSerialNumber(GetPayrollStartDate(lForPayrollID)))
				sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				Do While Not oRecordset.EOF
					sContents = Replace(sContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptShortName").Value) & " />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value)))
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop

				iStartPos = InStr(1, sContents, "<CONCEPT_", vbBinaryCompare)
				Do While (iStartPos > 0)
					iEndPos = InStr(iStartPos, sContents, "/>", vbBinaryCompare) + Len("/>")
					If iEndPos > 0 Then
						sContents = Left(sContents, (iStartPos - Len("<"))) & "&nbsp;" & Right(sContents, (Len(sContents) - iEndPos + Len(".")))
					End If
					iStartPos = InStr(1, sContents, "<CONCEPT_", vbBinaryCompare)
					If Err.Number <> 0 Then Exit Do
				Loop
				Response.Write sContents
			Else
				lErrorNumber = -1
				sErrorDescription = "La plantilla para armar el reporte está vacía."
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "No existen montos o pagos registrados para el empleado en la quincena especificada."
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	BuildReport1499 = lErrorNumber
	Err.Clear
End Function
%>