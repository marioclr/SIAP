
<%
Const N_SOCIETY_ID = 0
Const N_COMPANY_ID = 1
Const N_PERIOD_ID_FOR_SAR = 2
Const S_CLC = 3
Const N_BANK_ID = 4
Const S_BANK_SHORT_NAME = 5
Const N_PAYMENT_DATE = 6
Const N_EMPLOYEE_TYPE_ID = 7
Const S_EMPLOYEE_TYPE_NAME = 8
Const N_INCOME = 9
Const N_DEDUCTIONS = 10
Const N_NET_INCOME = 11
Const N_CPT_01 = 12
Const N_CPT_04 = 13
Const N_CPT_05 = 14
Const N_CPT_06 = 15
Const N_CPT_07 = 16
Const N_CPT_08 = 17
Const N_CPT_11 = 18
Const N_CPT_44 = 19
Const N_CPT_B2 = 20
Const N_CPT_7S = 21
Const N_USER_ID_FOR_SAR_FOR_SAR = 22
Const N_LAST_Update_DATE_FOR_SAR = 23
Const S_COMMENTS_FOR_SAR = 24
Const B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR = 25
Const B_IS_DUPLICATED_PAYROLL_RESUME_FOR_SAR = 26
Const B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT = 27

Const N_PAYROLL_RESUME_FOR_SAR_COMPONENT_SIZE = 27

Dim aPayrollResumeForSarComponent()
Redim aPayrollResumeForSarComponent(N_PAYROLL_FOR_SAR_COMPONENT_SIZE)

Function InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
'************************************************************
'Purpose: To initialize the empty elements of the 
'		  PayrollResumeForSar Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aPayrollResumeForSarComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePayrollResumeForSarComponent"
	Dim iItem
	Redim Preserve aPayrollResumeForSarComponent(N_PAYROLL_RESUME_FOR_SAR_COMPONENT_SIZE)

	aPayrollResumeForSarComponent(N_SOCIETY_ID) = -1
	aPayrollResumeForSarComponent(N_COMPANY_ID) = -1
	aPayrollResumeForSarComponent(N_PERIOD_ID_FOR_SAR) = -1
	aPayrollResumeForSarComponent(S_CLC) = ""
	aPayrollResumeForSarComponent(N_BANK_ID) = -1
	aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) = ""
	aPayrollResumeForSarComponent(N_PAYMENT_DATE) = -1
	aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = -1
	aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) = ""
	aPayrollResumeForSarComponent(N_INCOME) = 0
	aPayrollResumeForSarComponent(N_DEDUCTIONS) = 0
	aPayrollResumeForSarComponent(N_NET_INCOME) = 0
	aPayrollResumeForSarComponent(N_CPT_01) = 0
	aPayrollResumeForSarComponent(N_CPT_04) = 0
	aPayrollResumeForSarComponent(N_CPT_05) = 0
	aPayrollResumeForSarComponent(N_CPT_06) = 0
	aPayrollResumeForSarComponent(N_CPT_07) = 0
	aPayrollResumeForSarComponent(N_CPT_08) = 0
	aPayrollResumeForSarComponent(N_CPT_11) = 0
	aPayrollResumeForSarComponent(N_CPT_44) = 0
	aPayrollResumeForSarComponent(N_CPT_B2) = 0
	aPayrollResumeForSarComponent(N_CPT_7S) = 0
	aPayrollResumeForSarComponent(N_USER_ID_FOR_SAR) = 0
	aPayrollResumeForSarComponent(N_LAST_Update_DATE_FOR_SAR) = -1
	aPayrollResumeForSarComponent(S_COMMENTS_FOR_SAR) = ""


	aPayrollResumeForSarComponent(B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR) = True
	aPayrollResumeForSarComponent(B_IS_DUPLICATED_PAYROLL_RESUME_FOR_SAR) = False

	aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT) = True
	InitializePayrollResumeForSarComponent = Err.number
	Err.Clear
End Function

Function AddPayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new record into the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPayrollResumeForSarRecord"
	Dim lErrorNumber
	Dim bComponentInitialized
	bComponentInitialized = aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
	End If
	If lErrorNumber = 0 Then
		If aPayrollResumeForSarComponent(B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR) Then
			lErrorNumber = CheckExistencyOfPayrollResumeForSarRecord(aPayrollResumeForSarComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aPayrollResumeForSarComponent(B_IS_DUPLICATED_PAYROLL_RESUME_FOR_SAR) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Existen duplicidades en el registro (Empresa: " & aPayrollResumeForSarComponent(N_COMPANY_ID) & ",Banco: " & aPayrollResumeForSarComponent(N_BANK_ID) & ",Fecha de pago: " & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & ")"
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckPayrollResumeForSarInformationConsistency(aPayrollResumeForSarComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					lErrorNumber GetPayrollResumeForSarRecord(oRequest,oADODBConnection,aPayrollResumeForSarComponent,sErrorDescription)
					sErrorDescription = "No se pudo guardar la información del nuevo registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert into DM_HIST_NOMSAR (SocietyID,CompanyID,periodID,CLC,BankID,PaymentDate,EmployeeType,Income,Deductions,NetIncome,Cpt_01,Cpt_04,Cpt_05,Cpt_06,Cpt_07,Cpt_08,Cpt_11,Cpt_44,Cpt_b2,Cpt_7s,UserID,LastUpdateDate,Comments) Values (" & aPayrollResumeForSarComponent(N_SOCIETY_ID) & "," & aPayrollResumeForSarComponent(N_COMPANY_ID) & "," & aPayrollResumeForSarComponent(N_PERIOD_ID_FOR_SAR) & "," & aPayrollResumeForSarComponent(S_CLC) & "," & aPayrollResumeForSarComponent(N_BANK_ID) & "," & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & "," & aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) & "," & aPayrollResumeForSarComponent(N_INCOME) & "," & aPayrollResumeForSarComponent(N_DEDUCTIONS) & "," & aPayrollResumeForSarComponent(N_NET_INCOME) & "," & aPayrollResumeForSarComponent(N_CPT_01) & "," & aPayrollResumeForSarComponent(N_CPT_04) & "," & aPayrollResumeForSarComponent(N_CPT_05) & "," & aPayrollResumeForSarComponent(N_CPT_06) & "," & aPayrollResumeForSarComponent(N_CPT_07) & "," & aPayrollResumeForSarComponent(N_CPT_08) & "," & aPayrollResumeForSarComponent(N_CPT_11) & "," & aPayrollResumeForSarComponent(N_CPT_44) & "," & aPayrollResumeForSarComponent(N_CPT_B2) & "," & aPayrollResumeForSarComponent(N_CPT_7S) & "," & aPayrollResumeForSarComponent(N_USER_ID_FOR_SAR) & "," & aPayrollResumeForSarComponent(N_LAST_Update_DATE_FOR_SAR) & "," & aPayrollResumeForSarComponent(S_COMMENTS_FOR_SAR) & ")", "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If
	AddPayrollResumeForSarRecord = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeTypeIDFromShortName(oRequest)
'************************************************************
'Purpose: To get the EmployeeTypeID from the short name
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "GetEmployeeTypeIDFromShortName"
	Dim sQuery
	Dim oRecordsetEmployeeType
	Dim lErrorNumber

	sQuery = "Select EmployeeTypeID From EmployeeTypes Where (EmployeeTypeShortName = '" & oRequest("EmployeeTypeShortName").Item & "')"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetEmployeeType)
	GetEmployeeTypeIDFromShortName = oRecordsetEmployeeType.Fields("EmployeeTypeID").Value
	Set oRecordsetEmployeeType = Nothing
End Function

Function GetBankIDFromShortName(oRequest)
'************************************************************
'Purpose: To get the bankid from the short name
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "GetBankIDFromShortName"
	Dim sQuery
	Dim oRecordsetBank
	Dim lErrorNumber
	
	sQuery = "Select BankID From Banks Where (BankShortName = '" & oRequest("BankShortName").Item & "') Or (BankName = '" & oRequest("BankShortName").Item & "')"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetBank)
	GetBankIDFromShortName = oRecordsetBank.Fields("BankID").Value
	Set oRecordsetBank = Nothing
End Function

Function GetPayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a record from the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollResumeForSarRecord"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
	End If

	If aPayrollResumeForSarComponent(N_SOCIETY_ID) = -1 Or aPayrollResumeForSarComponent(N_COMPANY_ID) = -1 Or _
		aPayrollResumeForSarComponent(N_BANK_ID) = -1 Or aPayrollResumeForSarComponent(N_PAYMENT_DATE) = -1 Or _
		aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = -1 Or aPayrollResumeForSarComponent(S_CLC) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "La información proporcionada no permite ubicar un registro en el resumen de nóminas."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "aPayrollResumeForSarComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		sQuery = "Select * From DM_HIST_NOMSAR Where (SocietyID=" & aPayrollResumeForSarComponent(N_SOCIETY_ID) & ") And (CompanyID=" & aPayrollResumeForSarComponent(N_COMPANY_ID) & ") And (PaymentDate=" & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & ") And (BankID='" & aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) & "') And (EmployeeType='" & aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) & "') And (CLC=" & aPayrollResumeForSarComponent(S_CLC) & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aPayrollResumeForSarComponent(N_SOCIETY_ID) = oRecordset.Fields("SocietyID").Value
				aPayrollResumeForSarComponent(N_COMPANY_ID) = oRecordset.Fields("CompanyID").Value
				aPayrollResumeForSarComponent(N_PERIOD_ID_FOR_SAR) = oRecordset.Fields("PeriodID").Value
				aPayrollResumeForSarComponent(S_CLC) = oRecordset.Fields("CLC").Value
				aPayrollResumeForSarComponent(N_BANK_ID) = oRecordset.Fields("BankID").Value
				aPayrollResumeForSarComponent(N_PAYMENT_DATE) = oRecordset.Fields("PaymentDate").Value
				aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) = oRecordset.Fields("EmployeeType").Value
				aPayrollResumeForSarComponent(N_INCOME) = oRecordset.Fields("Income").Value
				aPayrollResumeForSarComponent(N_DEDUCTIONS) = oRecordset.Fields("Deductions").Value
				aPayrollResumeForSarComponent(N_NET_INCOME) = oRecordset.Fields("NetIncome").Value
				aPayrollResumeForSarComponent(N_CPT_01) = oRecordset.Fields("Cpt_01").Value
				aPayrollResumeForSarComponent(N_CPT_04) = oRecordset.Fields("Cpt_04").Value
				aPayrollResumeForSarComponent(N_CPT_05) = oRecordset.Fields("Cpt_05").Value
				aPayrollResumeForSarComponent(N_CPT_06) = oRecordset.Fields("Cpt_06").Value
				aPayrollResumeForSarComponent(N_CPT_07) = oRecordset.Fields("Cpt_07").Value
				aPayrollResumeForSarComponent(N_CPT_08) = oRecordset.Fields("Cpt_08").Value
				aPayrollResumeForSarComponent(N_CPT_11) = oRecordset.Fields("Cpt_11").Value
				aPayrollResumeForSarComponent(N_CPT_44) = oRecordset.Fields("Cpt_44").Value
				aPayrollResumeForSarComponent(N_CPT_B2) = oRecordset.Fields("Cpt_b2").Value
				aPayrollResumeForSarComponent(N_CPT_7S) = oRecordset.Fields("Cpt_7s").Value
				aPayrollResumeForSarComponent(N_USER_ID_FOR_SAR) = oRecordset.Fields("UserID").Value
				aPayrollResumeForSarComponent(N_LAST_Update_DATE_FOR_SAR) = oRecordset.Fields("LastUpdateDate").Value
				aPayrollResumeForSarComponent(S_COMMENTS_FOR_SAR) = oRecordset.Fields("Comments").Value
			End If
			oRecordset.Close
		End If
	End If
	Set oRecordset = Nothing
	GetPayrollResumeForSarRecord = lErrorNumber
	Err.Clear
End Function

Function GetPayrollResumeForSarList(oRequest, oADODBConnection, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about the proessional risk table
'Inputs:  oRequest, oADODBConnection
'Outputs: oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollResumeForSarList"
	Dim lErrorNumber
	Dim sQuery
	Dim lPeriodID
	Dim lPayrollID
	Dim asBanks
	Dim asCLCs
	Dim asCompanies
	Dim asClcCompanies
	Dim asConceptsAmount
	Dim asEmployeeTypes
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim lIndex
	Dim mIndex

	'Busca periodo abierto
	sQuery = "Select PeriodName From dm_sar_periods Where (IsOpen = 1)"
	sErrorDescription = "No se pudo obtener la información de los periodos abiertos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se encontraron periodos abiertos"
		lErrorNumber = -1
		If Not oRecordset.EOF Then
			'Obtiene el periodo actual
			lPeriodID = oRecordset.Fields("PeriodName").Value
			lErrorNumber = 0
			'Obtención de CLCs relacionadas al periodo abierto
            sQuery = "Select Distinct PayrollCLC From PayrollsCLCs Where (PayrollCode = '" & lPeriodID & "') Order By 1 Asc"
			sErrorDescription = "No se pudo obtener la infomración de las CLCs"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se encontraron CLCs relacionadas al periodo " & lPeriodID
				lErrorNumber = -1
				If Not oRecordset.EOF Then
					lErrorNumber = 0
					asCLCs = oRecordset.GetRows
					sQuery = "Select CompanyID, CompanyName From Companies Order By CompanyID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					asCompanies = oRecordset.GetRows
					sQuery = "Delete From DM_Hist_Nomsar Where (CLC Not Like 'x%')"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					For iIndex = 0 To UBound(asCLCs,2)
						'Extracción de Payroll relacionado a la CLC
						sQuery = "Select Distinct PayrollID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If Not oRecordset.EOF Then
							lPayrollID = oRecordset.Fields("PayrollID").Value
							'Se obtienen las empresas relacionadas a la CLC
							sQuery = "Select Distinct CompanyID From EmployeesHistoryListForPayroll EHL, (Select EmployeeID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')) CLC Where (EHL.EmployeeID = CLC.EmployeeID) And (EHL.PayrollID In (" & lPayrollID & "))"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							asClcCompanies = oRecordset.GetRows
							'Se obtienen los bancos relacionados a la CLC
							sQuery = "Select Distinct BankID From EmployeesHistoryListForPayroll EHL, (Select EmployeeID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')) CLC Where (EHL.EmployeeID = CLC.EmployeeID) And (EHL.PayrollID In (" & lPayrollID & "))"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							asBanks = oRecordset.GetRows
							'Se obtienen los tipos de empleado relacionados a la CLC
							sQuery = "Select Distinct EmployeeTypeID From EmployeesHistoryListForPayroll EHL, (Select EmployeeID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')) CLC Where (EHL.EmployeeID = CLC.EmployeeID) And (EHL.PayrollID In (" & lPayrollID & "))"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							asEmployeeTypes = oRecordset.GetRows
							'Cálculo de conceptos en la nómina relacionada a la CLC por Empresa
							For jIndex = 0 To UBound(asClcCompanies,2)
								For kIndex = 0 To UBound(asBanks,2)
									For lIndex = 0 To UBound(asEmployeeTypes,2)
										sQuery = "Select 1, CompanyID, " & lPeriodID & ", " & asCLCs(0,iIndex) & ", BankID, " & lPayrollID & ", EmployeeTypeID, ConceptID, Sum(ConceptAmount) Importe From Payroll_" & Mid(lPeriodID,1,4) & " Pr, EmployeesHistoryListForPayroll EHL, (Select EmployeeID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')) CLC Where (Pr.EmployeeID = CLC.EmployeeID) And (EHL.EmployeeID = Pr.EmployeeID) And (EHL.PayrollID In (" & lPayrollID & ")) And (EHL.CompanyID = " & asClcCompanies(0,jIndex) & ") And (EHL.BankID = " & asBanks(0,kIndex) & ") And (EHL.EmployeeTypeID = " & asEmployeeTypes(0,lIndex) & ") And (Pr.RecordID In(" & lPayrollID & ")) And (Pr.ConceptID In (-2,-1,0,1,4,5,6,7,8,13,47,89,146)) Group By CompanyID, BankID, EmployeeTypeID, ConceptID Order By CompanyID, BankID, EmployeeTypeID, ConceptID Asc"
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												asConceptsAmount = oRecordset.getRows()
												sQuery = "Insert Into DM_Hist_Nomsar Values (1," & asClcCompanies(0,jIndex) & "," & _
													lPeriodID & ",'" & asCLCs(0,iIndex) & "','" & asBanks(0,kIndex) & "'," & lPayrollID & ",'" & _
													asEmployeeTypes(0,lIndex) & "', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0," & _
													aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ", '')"
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												sQuery = "Update DM_Hist_Nomsar Set "
												For mIndex = 0 To UBound(asConceptsAmount,2)
													If CLng(asConceptsAmount(7,mIndex)) = -2 Then sQuery = sQuery & "|Deductions = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = -1 Then sQuery = sQuery & "|Income = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 0 Then sQuery = sQuery & "|NetIncome = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 1 Then sQuery = sQuery & "|Cpt_01 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 4 Then sQuery = sQuery & "|Cpt_04 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 5 Then sQuery = sQuery & "|Cpt_05 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 6 Then sQuery = sQuery & "|Cpt_06 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 7 Then sQuery = sQuery & "|Cpt_07 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 8 Then sQuery = sQuery & "|Cpt_08 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 13 Then sQuery = sQuery & "|Cpt_11 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 47 Then sQuery = sQuery & "|Cpt_44 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 89 Then sQuery = sQuery & "|Cpt_B2 = " & asConceptsAmount(8,mIndex) & "|"
													If CLng(asConceptsAmount(7,mIndex)) = 146 Then sQuery = sQuery & "|Cpt_7S = " & asConceptsAmount(8,mIndex) & "|"
												Next
												sQuery = Replace(sQuery,"||",",")
												sQuery = Replace(sQuery,"|","")
												sQuery = sQuery & " Where (CompanyID = " & asClcCompanies(0,jIndex) & ") And (PeriodID = " & lPeriodID & ") " & _
														"And (CLC = '" & asCLCs(0,iIndex) & "') And (BankID = " & asBanks(0,kIndex) & ") " & _
														"And (EmployeeType = " & asEmployeeTypes(0,lIndex) & ")"
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
										End If
									Next
								Next
							Next
						End If
					Next
					sQuery = "Select * From DM_Hist_Nomsar Where (CLC Not Like 'x%') And (PeriodID ="  & lPeriodID & ") Order By PaymentDate, CLC, CompanyID, BankID, EmployeeType Asc"
					sErrorDescription = "No se pudo leer la información para el resumen"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							sErrorDescription = "No existe información calculada del periodo " & lPeriodID
							lErrorNumber = -1
						End If
					End If
				End If
			End If
		End If
	End If
	GetPayrollResumeForSarList = lErrorNumber
	Err.Clear
End Function

Function ModifyPayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing record in the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPayrollResumeForSarRecord"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
	End If

	If aPayrollResumeForSarComponent(B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR) Then
		lErrorNumber = CheckExistencyOfPayrollResumeForSarRecord(aPayrollResumeForSarComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If Not CheckPayrollResumeForsarInformationConsistency(aPayrollResumeForSarComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update DM_HIST_NOMSAR Set Income=" & aPayrollResumeForSarComponent(N_INCOME) & ",Deductions=" & aPayrollResumeForSarComponent(N_DEDUCTIONS) & ",NetIncome=" & aPayrollResumeForSarComponent(N_NET_INCOME) & ",Cpt_01=" & aPayrollResumeForSarComponent(N_CPT_01) & ",Cpt_04=" & aPayrollResumeForSarComponent(N_CPT_04) & ",Cpt_05=" & aPayrollResumeForSarComponent(N_CPT_05) & ",Cpt_06=" & aPayrollResumeForSarComponent(N_CPT_06) & ",Cpt_07=" & aPayrollResumeForSarComponent(N_CPT_07) & ",Cpt_08=" & aPayrollResumeForSarComponent(N_CPT_08) & ",Cpt_11=" & aPayrollResumeForSarComponent(N_CPT_11) & ",Cpt_44=" & aPayrollResumeForSarComponent(N_CPT_44) & ",Cpt_b2=" & aPayrollResumeForSarComponent(N_CPT_B2) & ",Cpt_7s=" & aPayrollResumeForSarComponent(N_CPT_7S) & " Where (SocietyId=" & aPayrollResumeForSarComponent(N_SOCIETY_ID) & ") And (CompanyID= " & aPayrollResumeForSarComponent(N_COMPANY_ID) & ") And (BankID='" & aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) & "') And (PaymentDate=" & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & ") And (EmployeeType='" & aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) & "') And CLC='" & aPayrollResumeForSarComponent(S_CLC) & "'", "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If
	ModifyPayrollResumeForSarRecord = lErrorNumber
	Err.Clear
End Function

Function RemovePayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a record from the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePayrollResumeForSarRecord"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
	End If

	If aPayrollResumeForSarComponent(B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR) Then
		lErrorNumber = CheckExistencyOfPayrollResumeForSarRecord(aPayrollResumeForSarComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If Not CheckPayrollResumeForSarInformationConsistency(aPayrollResumeForSarComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From DM_Hist_Nomsar Where (SocietyId=" & aPayrollResumeForSarComponent(N_SOCIETY_ID) & ") And (CompanyID= " & aPayrollResumeForSarComponent(N_COMPANY_ID) & ") And (BankID='" & aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) & "') And (PaymentDate=" & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & ") And (EmployeeType='" & aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) & "') And CLC='" & aPayrollResumeForSarComponent(S_CLC) & "'", "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemovePayrollResumeForSarRecord = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPayrollResumeForSarRecord(aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record exists in the table
'Inputs:  aPayrollResumeForSarComponent
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPayrollResumeForSarRecord"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollResumeforSarComponent(oRequest, aPayrollResumeForSarComponent)
	End If

	If aPayrollResumeForSarComponent(N_SOCIETY_ID) = -1 Or aPayrollResumeForSarComponent(N_COMPANY_ID) = -1 Or aPayrollResumeForSarComponent(N_BANK_ID) = -1 Or aPayrollResumeForSarComponent(N_PAYMENT_DATE) = -1 Or aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "La información proporcionada no permite ubicar un registro en el resumen de nóminas."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From DM_HIST_NOMSAR Where (SocietyID=" & oRequest("SocietyID").Item & ") And (CompanyID=" & oRequest("CompanyID").Item & ") And (CLS=" & oRequest("CLC").Item & ")And (PaymentDate=" & oRequest("PaymentDate").Item & ") And (BankID=" & oRequest("BankID").Item & ") And (EmployeeTypeID=" & oRequest("EmployeeTypeID").Item & ")", "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aPayrollResumeForSarComponent(B_IS_DUPLICATED_PAYROLL_RESUME_FOR_SAR) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPayrollResumeForSarRecord = lErrorNumber
	Err.Clear
End Function

Function CheckPayrollResumeForSarInformationConsistency(aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the table
'Inputs:  aPayrollResumeForSarComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckPayrollResumeForSarInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aPayrollResumeForSarComponent(N_COMPANY_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El inicador de la empresa no es numérico."
		bIsCorrect = False
	End If
	
	If Not IsNumeric(aPayrollResumeForSarComponent(N_PERIOD_ID_FOR_SAR)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;-	El campo del periodo no es numérico."
		bIsCorrect = False
	End If

	If Not IsNumeric(aPayrollResumeForSarComponent(N_BANK_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El indicador del banco no es un valor numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_PAYMENT_DATE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de pago no es un valor numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El campo con el tipo de empleado debe ser numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_INCOME)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El campo Ingresos debe ser numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_DEDUCTIONS)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El campo Deducciones debe ser numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_NET_INCOME)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El campo Líquido debe ser numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_01)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 01 debe ser numérico."
		bIsCorrect = False
	End If

	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_04)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 04 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_05)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 05 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_06)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 06 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_07)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 07 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_08)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 08 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_11)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 11 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_44)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 44 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_B2)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto B2 debe ser numérico."
		bIsCorrect = False
	End If
	
	If Not isNumeric(aPayrollResumeForSarComponent(N_CPT_7S)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del concepto 7S debe ser numérico."
		bIsCorrect = False
	End If
	CheckPayrollResumeForSarInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayPeriodsList(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the list of periods
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPeriodsList"
	Dim sNames
	Dim sCondition
	Dim oRecordset
	Dim sQuery
	Dim lErrorNumber
	Dim iStartPage
	Dim sRowContents
	Dim asRowContents
	Dim asColumnsTitles
	Dim asCellWidths
	Dim asCellAlignments
	Dim asTableColors()
	Dim iRecordCounter
		
	sQuery = "Select * From DM_SAR_PERIODS Order By PeriodID Desc"
	sErrorDescription = "No se pudieron obtener los periodos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se encontraron periodos registrados"
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Periodo, De nómina, A Nómina, Estatus", ",", -1, vbBinaryCompare)
			asCellWidths = Split("100,100,100,50", ",", -1, vbBinaryCompare)
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
			asCellAlignments = Split("CENTER,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = ""
				sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PeriodName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StartPayroll").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EndPayroll").Value))
				If CInt(oRecordset.Fields("IsOpen").Value) = 1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Abierta")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Cerrada")
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				oRecordset.MoveNext
				iRecordCounter = iRecordCounter + 1
				If (iRecordCounter >= ROWS_CATALOG) Then Exit Do
				If Err.number <> 0 Then Exit Do
			Loop
			Response.Write "</TABLE>" & vbNewLine
		End If
	End If

	DisplayPeriodsList = lErrorNumber
	Err.Clear
End Function


Function DisplayConsarFile(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the table in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aPayrollResumeForSarComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConsarFile"
	Dim sNames
	Dim sCondition
	Dim oRecordset
	Dim sQuery
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	Dim iStartPage
	Dim iRecordCounter	

	sCondition = ""
	If Len(oRequest("EmployeeNumber").Item) > 0 Then 
		If Len(oRequest("EmployeeNumber").Item) = 6 Then
			sCondition = sCondition & " EmployeeID Like '" & CLng(oRequest("EmployeeNumber").Item) & "'|"  
		Else
			sCondition = sCondition & " EmployeeID Like '%" & oRequest("EmployeeNumber").Item & "%'|" 
		End If
	End If
	If Len(oRequest("EmployeeName").Item) > 0 Then sCondition = sCondition & " EmployeeName Like '% " & oRequest("EmployeeName").Item & "%'|"
	If Len(oRequest("EmployeeLastName").Item) > 0 Then sCondition = sCondition & " EmployeeLastName Like '%" & oRequest("EmployeeLastName").Item & "%'|"
	If Len(oRequest("EmployeeLastName2").Item) > 0 Then sCondition = sCondition & " EmployeeLastName2 Like '%" & oRequest("EmployeeLastName2").Item & "%'|"
	If Len(oRequest("RFC").Item) > 0 Then sCondition = sCondition & " RFC Like '%" & oRequest("RFC").Item & "%'|"
	If Len(oRequest("CURP").Item) > 0 Then sCondition = sCondition & " CURP Like '%" & oRequest("CURP").Item & "%'|"
	If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " u_version = " & oRequest("CompanyID").Item & "|"
	
	If Len(sCondiction) > 0 Then
		sCondition = " Where " & Replace(Mid(sCondition,1,Len(sCondition)-1),"|"," And ")
	End If
	
	iRecordCounter = 1
	sQuery = "Select * From DM_APORT_SAR " & sCondition
	sErrorDescription = "No se pudo obteener la información del archivo de la CONSAR."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Clave, Filiación, CURP, NSS, Apellido paterno, Apellido materno, Nombre, Nombramiento, Clave ICEFA, Fecha de ingreso, Fecha de cotización, Fovissste, Días cotizados, Días incapacidad, Días ausencias, Salario base, Salario Base V, SAR, CV Patrón, CV Trabajador, FOVI, Ahorro trabajador, Ahorro dependencia", ",", -1, vbBinaryCompare)
			asCellWidths = Split("50,100,100,100,200,200,200,100,100,200,200,100,50,50,50,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,,,,,,,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ProfileID"" ID=""ProfileIDRd"" VALUE=""" & CStr(oRecordset.Fields("ProfileID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ProfileID"" ID=""ProfileIDChk"" VALUE=""" & CStr(oRecordset.Fields("ProfileID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("cve").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("rfc").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("curp").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & " "
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("nombram").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("icefa").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("JoinDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("cotDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("fovi").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("workingDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("inabilityDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("absenceDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("salary").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("salaryV").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("sar").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("entityCV").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("employeeCV").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("foviAmount").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("employeeSaving").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("entitySaving").Value, 2, True, False, True)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (iRecordCounter >= ROWS_CATALOG) Then Exit Do
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = -1
			sErrorDescription = "El archivo no se ha cargado o está vacío."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayConsarFile = lErrorNumber
	Err.Clear

End Function

Function DisplayStartSarForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the fields for start a new period
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayStartSarForm"
	Dim lPeriod
	Dim asPeriods
	Dim iIndex
	
	asPeriods=Split("1,1,2,2,3,3,4,4,5,5,6,6",",")
	lPeriod = asPeriods(Month(Date)-1)
	
	Call DisplayInstructionsMessage("Información", "Para abrir el nuevo bimestre seleccione ls nóminas inicial y final que acotarán el periodo. <BR />CuAndo esté listo, presione el botón <B>Agregar</B> para continuar.")
	
	Response.Write "<FORM NAME=""PayrollFrm"" ID=""PayrollFrm"" ACTION=""PayrollResumeForSar.asp"" METHOD=""POST"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" VALUE=""StartPeriod"" NAME=""Action"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" VALUE=""2"" NAME=""Step"">"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año:&nbsp;</FONT></TD>"
				Response.Write "&nbsp;&nbsp;<TD><Select NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = (Year(Date()) - 2) To Year(Date()) + 1
						Response.Write "<OPTION VALUE=""" & iIndex
						If Year(Date()) = iIndex Then Response.Write """ SelectED=""1"""
						Response.Write  """>" & iIndex & "</OPTION>"
					Next
				Response.Write "</Select></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"">&nbsp;</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Bimestre:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""1"" "
						If lPeriod = 1 Then Response.Write " CHECKED=""1"""
					Response.Write " />01"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""2"""
						If lPeriod = 2 Then Response.Write " CHECKED=""1"""
					Response.Write " />02"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""3"""
						If lPeriod = 3 Then Response.Write " CHECKED=""1"""
					Response.Write " />03"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""4"""
						If lPeriod = 4 Then Response.Write " CHECKED=""1"""
					Response.Write " />04"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""5"""
						If lPeriod = 5 Then Response.Write " CHECKED=""1"""
					Response.Write " />05"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodID"" VALUE=""6"""
						If lPeriod = 6 Then Response.Write " CHECKED=""1"""
					Response.Write " />06"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
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
	Response.Write "</FORM>"
	DisplayStartSarForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollResumeForSarForm(oRequest, oADODBConnection, sAction, aPayrollResumeForSarComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record from the
'		  table using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aPayrollResumeForSarComponent
'Outputs: aPayrollResumeForSarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollResumeForSarForm"
	Dim sNames
	Dim sTempNames
	Dim lErrorNumber
	Dim sPosition
	Dim sService
		If aPayrollResumeForSarComponent(N_SOCIETY_ID) = -1 Or aPayrollResumeForSarComponent(N_COMPANY_ID) = -1 Or _
			aPayrollResumeForSarComponent(N_BANK_ID) = -1 Or aPayrollResumeForSarComponent(N_PAYMENT_DATE) = -1 Or _
			aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = -1 Or aPayrollResumeForSarComponent(S_CLC) = "" Then
			aPayrollResumeForSarComponent(N_SOCIETY_ID) = oRequest("SocietyID").Item
			aPayrollResumeForSarComponent(N_COMPANY_ID) = oRequest("CompanyID").Item
			aPayrollResumeForSarComponent(N_BANK_ID) = GetBankIDFromShortName(oRequest)
			aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) = oRequest("BankShortName").Item
			aPayrollResumeForSarComponent(N_PAYMENT_DATE) = oRequest("PaymentDate").Item
			aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = GetEmployeeTypeIDFromShortName(oRequest)
			aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) = oRequest("EmployeeTypeShortName").Item
			aPayrollResumeForSarComponent(S_CLC) = oRequest("CLC").Item
			aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT) = True
			lErrorNumber = GetPayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function CheckConceptFields(oForm) {" & vbNewLine
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "oForm.Income.value = oForm.Income.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Income, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Deductions.value = oForm.Deductions.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Deductions, 'El monto de las deducciones', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.NetIncome.value = oForm.NetIncome.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.NetIncome, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_01.value = oForm.Cpt_01.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_01, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_04.value = oForm.Cpt_04.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_04, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_05.value = oForm.Cpt_05.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_05, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_06.value = oForm.Cpt_06.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_06, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_07.value = oForm.Cpt_07.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_07, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_08.value = oForm.Cpt_08.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_08, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_11.value = oForm.Cpt_11.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_11, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_44.value = oForm.Cpt_44.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_44, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_7S.value = oForm.Cpt_7S.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_7S, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.Cpt_B2.value = oForm.Cpt_B2.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.Cpt_B2, 'El monto de ingresos brutos', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				Response.Write "<FORM NAME=""PayrollResumeForSarFrm"" ID=""PayrollResumeForSarFrm"" ACTION=""Catalogs.asp"" METHOD=""POST"" "
				If Len(oRequest("Modify").Item) > 0 Then Response.Write " onSubmit=""return CheckConceptFields(this)"" "
				Response.Write ">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PayrollResume"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SocietyID"" ID=""SocietyIDHdn"" VALUE=""" & aPayrollResumeForSarComponent(N_SOCIETY_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aPayrollResumeForSarComponent(N_COMPANY_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BankID"" ID=""BankIDHdn"" VALUE=""" & aPayrollResumeForSarComponent(N_BANK_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentDate"" ID=""PaymentDateHdn"" VALUE=""" & aPayrollResumeForSarComponent(N_PAYMENT_DATE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeType"" ID=""EmployeeTypeHdn"" VALUE=""" & aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""clc"" ID=""clcHdn"" VALUE=""" & aPayrollResumeForSarComponent(S_CLC) & """ />"
				Response.Write "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sociedad:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aPayrollResumeForSarComponent(S_SOCIETY_ID) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Companies", aPayrollResumeForSarComponent(N_COMPANY_ID), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aPayrollResumeForSarComponent(N_PERIOD_ID_FOR_SAR) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CLC:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aPayrollResumeForSarComponent(S_CLC) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Banco:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de pago:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(CStr(aPayrollResumeForSarComponent(N_PAYMENT_DATE)), -1, -1, -1) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de empleado:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
					Response.Write "</TR>"
					If Len(oRequest("Modify").Item) > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ingresos:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Income"" ID=""IncomeTxt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_INCOME), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Deducciones:</FONT></TD><TD ALIGN=""RIGHT"" ><INPUT TYPE=""TEXT"" NAME=""Deductions"" ID=""DeductionsTxt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_DEDUCTIONS), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Líquido:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""NetIncome"" ID=""NetIncomeTxt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_NET_INCOME), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""20"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_01:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_01"" ID=""Cpt_01Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_01), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_04:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_04"" ID=""Cpt_04Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_04), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_05:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_05"" ID=""Cpt_05Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_05), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_06:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_06"" ID=""Cpt_06Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_06), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_07:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_07"" ID=""Cpt_07Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_07), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_08:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_08"" ID=""Cpt_08Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_08), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_11:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_11"" ID=""Cpt_11Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_11), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_44:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_44"" ID=""Cpt_44Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_44), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_7S:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_7S"" ID=""Cpt_7STxt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_7S), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_B2:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Cpt_B2"" ID=""Cpt_B2Txt"" VALUE=""" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_B2), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""15"" ALIGN=""RIGHT"" CLASS=""TextFields"" STYLE=""text-align: right;""/></TD>"
						Response.Write "</TR>"
					ElseIf Len(oRequest("Delete").Item > 0) Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ingresos:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_INCOME), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Deducciones:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_DEDUCTIONS), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Líquido:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_NET_INCOME), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_01:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_01), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_04:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_04), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_05:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_05), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_06:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_06), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_07:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_07), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_08:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_08), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_11:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_11), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_44:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_44), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_7S:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_7S), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cpt_B2:</FONT></TD><TD ALIGN=""RIGHT"" ><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aPayrollResumeForSarComponent(N_CPT_B2), 2, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
					End If
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR />"

				If aPayrollResumeForSarComponent(N_ID_PROFILE) = -1 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
				ElseIf Len(oRequest("Delete").Item) > 0 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Remove"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" />"
				Else
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ModifyResume"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				End If
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Catalogs.asp?Action=PayrollResume'"" />"
				Response.Write "<BR /><BR />"
				Call DisplayWarningDiv("RemoveCatalogWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
			Response.Write "</FORM>"
		End If
	DisplayPayrollResumeForSarForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollCompareList(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the table
'Inputs:  oRequest, oADODBConnection,
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollCompareList"
	Dim sNames
	Dim oRecordset
	Dim sQuery
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lCurrentPeriod
	Dim lErrorNumber
	Dim iStartPage
	Dim iRecordCounter	

    Dim sCompanyName

	iRecordCounter = 1
	sQuery = "Select PeriodName From Dm_Sar_Periods Where (IsOpen = 1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lCurrentPeriod = oRecordset.Fields("PeriodName").Value
	sQuery = "Select * From DM_Hist_Binmar Where (PeriodID = " & lCurrentPeriod & ") Order By SocietyID, CompanyID, PeriodID, PaymentDate"
	sErrorDescription = "No se pudo obteener la información del comparativo de devengos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Sociedad,Empresa,Periodo,Fecha de pago,Ingresos,Deducciones,Líquido,Cpt_n01,Cpt_n04,Cpt_n05,Cpt_n06,Cpt_n07,Cpt_n08,Cpt_n11,Cpt_n44,Cpt_nb2,Cpt_n7s,Cpt_a01,Cpt_a04,Cpt_a05,Cpt_a06,Cpt_a07,Cpt_a08,Cpt_a11,Cpt_a44,Cpt_ab2,Cpt_a7s,Diferencia,Val_01,Val_04,Val_05,Val_06,Val_07,Val_08,Val_11,Val_44,Val_b2,Val_7s,UserID,LastUpdateDate,Comments", ",", -1, vbBinaryCompare)
			asCellWidths = Split("50,100,100,200,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,150,100,150,250", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocietyID").Value))

                    Call GetNameFromTable(oADODBConnection, "Companies", oRecordset.Fields("CompanyID").Value, "", "", sCompanyName, sErrorDescription)
					'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyID").Value))
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(sCompanyName))

                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PeriodID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("PaymentDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Income").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Deductions").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("NetIncome").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n01").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n04").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n05").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n06").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n07").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n08").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n11").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n44").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_nb2").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_n7s").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a01").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a04").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a05").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a06").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a07").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a08").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a11").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a44").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_ab2").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_a7s").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					If (CInt(oRecordset.Fields("Val_01").Value) = 1) Or (IsNull(oRecordset.Fields("Val_01").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_04").Value) = 1) Or (IsNull(oRecordset.Fields("Val_04").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_05").Value) = 1) Or (IsNull(oRecordset.Fields("Val_05").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_06").Value) = 1) Or (IsNull(oRecordset.Fields("Val_06").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR &	"<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR &	"<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_07").Value) = 1) Or (IsNull(oRecordset.Fields("Val_07").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_08").Value) = 1) Or (IsNull(oRecordset.Fields("Val_08").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_11").Value) = 1) Or (IsNull(oRecordset.Fields("Val_11").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_44").Value) = 1) Or (IsNull(oRecordset.Fields("Val_44").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else 
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_b2").Value) = 1) Or (IsNull(oRecordset.Fields("Val_b2").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					If (CInt(oRecordset.Fields("Val_7s").Value) = 1) Or (IsNull(oRecordset.Fields("Val_7s").Value)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>F</FONT>"
					Else 
						sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """>V</FONT>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("LastUpdateDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (iRecordCounter >= ROWS_CATALOG) Then Exit Do
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = -1
			sErrorDescription = "No se ha generado el comparativo del periodo " & lCurrentPeriod
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayPayrollCompareList = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollResumeForSarList(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the table in a table
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollResumeForSarList"
	Dim sNames
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
	Dim iStartPage
	Dim iRecordCounter
	Dim sUserName
	Dim sCompanyName
	Dim sEmployeeTypeName

    Dim sBankName

	iRecordCounter = 1
	lErrorNumber = GetPayrollResumeForSarList(oRequest, oADODBConnection, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Sociedad,Empresa,Periodo,CLC,Banco,Fecha de pago,Tipo de empleado,Ingresos,Deducciones,Líquido,Cpt_01,Cpt_04,Cpt_05,Cpt_06,Cpt_07,Cpt_08,Cpt_11,Cpt_44,Cpt_b2,Cpt_7s,Usuario,Modificado,Comentarios,Acciones", ",", -1, vbBinaryCompare)
			asCellWidths = Split("50,100,250,50,50,50,250,100,100,100,100,100,100,100,100,100,100,100,100,100,250,250,250", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,,,,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					Call GetNameFromTable(oADODBConnection, "Users", oRecordset.Fields("UserID").Value, "", "", sUserName, sErrorDescription)
					Call GetNameFromTable(oADODBConnection, "Companies", oRecordset.Fields("CompanyID").Value, "", "", sCompanyName, sErrorDescription)
					Call GetNameFromTable(oADODBConnection, "EmployeeTypes", oRecordset.Fields("EmployeeType").Value, "", "", sEmployeeTypeName, sErrorDescription)
                    Call GetNameFromTable(oADODBConnection, "Banks", oRecordset.Fields("BankId").Value, "", "", sBankName, sErrorDescription)

					sRowContents = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocietyID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(sCompanyName))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PeriodID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CLC").Value))
					
                    'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BankID").Value))
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(sBankName))

					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("PaymentDate").Value), -1, -1, -1)
					
                    'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeType").Value))
                    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(sEmployeeTypeName))

					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Income").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Deductions").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("NetIncome").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_01").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_04").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_05").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_06").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_07").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_08").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_11").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_44").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_B2").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Cpt_7S").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sUserName)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("LastUpdateDate").Value), -1, -1, -1)
					If (Not IsNull(oRecordset.Fields("Comments").Value)) And (Len(oRecordset.Fields("Comments").Value) > 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					Else 
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("")
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("SocietyID").Value) <> 0 Then
							sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=PayrollResume&SocietyID=" & CStr(oRecordset.Fields("SocietyID").Value) & "&CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & "&CLC=" & oRecordset.Fields("CLC").Value & "&BankShortName=" & CStr(oRecordset.Fields("BankID").Value) & "&PaymentDate=" & CStr(oRecordset.Fields("PaymentDate").Value) & "&EmployeeTypeShortName=" & oRecordset.Fields("EmployeeType").Value & "&Modify=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_Delete_PERMISSIONS) = N_Delete_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=PayrollResume&SocietyID=" & CStr(oRecordset.Fields("SocietyID").Value) & "&CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & "&CLC=" & oRecordset.Fields("CLC").Value & "&BankShortName=" & CStr(oRecordset.Fields("BankID").Value) & "&PaymentDate=" & CStr(oRecordset.Fields("PaymentDate").Value) & "&EmployeeTypeShortName=" & oRecordset.Fields("EmployeeType").Value & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (iRecordCounter >= ROWS_CATALOG) Then Exit Do
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
	DisplayPayrollResumeForSarList = lErrorNumber
	Err.Clear
End Function

Function ComparePayrolls(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To search movements of employees between payrolls
'			And to define differences between employee's records
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ComparePayrolls"
	Dim sNames
	Dim oRecordset
	Dim lErrorNumber
	Dim lStartPayroll
	Dim lEndPayroll
	Dim lPeriodYear
	Dim lPeriod
	Dim sQuery
	Dim asPayrolls
	Dim iIndex
	Dim lTotal
	Dim lCurPeriod

	sQuery  = "Select PeriodName, StartPayroll, EndPayroll, PeriodYear From Dm_Sar_Periods Where (IsOpen = 1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lStartPayroll = oRecordset.Fields("StartPayroll").Value
	lEndPayroll = oRecordset.Fields("EndPayroll").Value
	lPeriod = oRecordset.Fields("PeriodName").Value
	lPeriodYear = oRecordset.Fields("PeriodYear").Value
	sQuery = "Select PayrollID From Payrolls Where PayrollID >= " & lStartPayroll & " And PayrollID <= " & lEndPayroll & " And PayrollTypeID = 1"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asPayrolls = oRecordset.GetRows()

	sQuery = "Insert Into Dm_Padron_Banamex (EmployeeID, Status, Period, ChangeFlag, JoinDate, CotDate, Nombram, ICNumber, mot_baja, WorkingDays, InabilityDays, AbsenceDays) " & _
				"Select Distinct EmployeeId, 2, " & asPayrolls(0,iIndex) & ", 'A'," & asPayrolls(0,iIndex) & ", " & asPayrolls(0,iIndex) & ",0,0,0," & 15*(4-iIndex) & ",0,0 " & _
				"From Dm_Estr_Qna" & _
				" Where (EmployeeID Not In (Select EmployeeId From Dm_Padron_Banamex)) " & _
				"Order By EmployeeID"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	
    If iConnectionType <> ORACLE Then
	sQuery = "Update Dm_Padron_Banamex Set Rfc = Employees.Rfc, Curp = Employees.Curp, " & _
				"SocialSecurityNumber = Employees.SocialSecurityNumber, EmployeeLastName = Employees.EmployeeLastName, " & _
				"EmployeeLastName2 = Employees.EmployeeLastName2, EmployeeName = Employees.EmployeeName, " & _
				"CT = Employees.PaymentCenterID, BirthDate = Employees.BirthDate, GenderShortName = Employees.GenderID, " & _
				"MaritalStatusID = Employees.MaritalStatusID " & _
			"From Employees " & _
			"Where (Dm_Padron_Banamex.Status = 2) " & _
				"And (Dm_Padron_Banamex.EmployeeID = Employees.EmployeeID)"

    Else
    sQuery = "Update Dm_Padron_Banamex Set (Rfc, Curp, SocialSecurityNumber, EmployeeLastName, EmployeeLastName2, EmployeeName, CT, BirthDate, GenderShortName,MaritalStatusID)=( " & _
             "Select b.Rfc,b.Curp,b.SocialSecurityNumber,b.EmployeeLastName,b.EmployeeLastName2,b.EmployeeName,b.PaymentCenterID,b.BirthDate, b.GenderID,b.MaritalStatusID " & _
             "From Employees b Where b.EmployeeID = Dm_Padron_Banamex.EmployeeID) " & _
             "Where (Dm_Padron_Banamex.status=2)"
    End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	
    If iConnectionType <> ORACLE Then
	sQuery = "Update Dm_Padron_Banamex Set Salary = Dm_Resultsar.Sar, Salary_v = Dm_Resultsar.fovi, FullPay  = (Dm_Resultsar.Tot01 + Dm_Resultsar.Tot06) " & _
			"From Dm_Resultsar Where (Dm_Padron_Banamex.EmployeeID = Dm_Resultsar.EmployeeID) And (Dm_Padron_Banamex.Status=2)"

    Else
    sQuery = "Update Dm_Padron_Banamex Set (Salary, Salary_v, FullPay)= " & _
			 "(Select Dm_Resultsar.Sar, Dm_Resultsar.fovi, Dm_Resultsar.Tot01 + Dm_Resultsar.Tot06  From Dm_Resultsar Where Dm_Padron_Banamex.EmployeeID = Dm_Resultsar.EmployeeID) " & _
             "Where (Dm_Padron_Banamex.Status=2)"
    End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

    sQuery="Select Count(*) Total From all_tables Where (table_name='DM_ESTR_QNA_TMP01') And (owner='SIAP')"
    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
    If oRecordset.Fields("Total").Value = 1 Then
       sQuery= "Truncate Table DM_ESTR_QNA_TMP01"
       lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
    Else
       sQuery= "Create Global Temporary Table DM_ESTR_QNA_TMP01( EMPLOYEEID  INTEGER PRIMARY KEY)"
       lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
    End If

    sQuery = "Insert Into DM_ESTR_QNA_TMP01 Select DISTINCT EmployeeID From Dm_Estr_Qna"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	'sQuery = "Update Dm_Padron_Banamex Set Status=3, mot_baja=7 Where (EmployeeId Not In (Select Distinct EmployeeID From Dm_Estr_Qna)) And (Status <> 2)"
    sQuery = "Update Dm_Padron_Banamex Set Status=3, mot_baja=7 Where EmployeeId Not In (Select EmployeeID From DM_ESTR_QNA_TMP01) And Status <> 2"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	sQuery = "Update Dm_resultsar Set Status = 'c' Where EmployeeID In (" & _
				"Select Dm_Resultsar.EmployeeID " & _
				"From Dm_Resultsar, Dm_Padron_Banamex  " & _
				"Where (Dm_Resultsar.EmployeeID = Dm_Padron_Banamex.EmployeeID) " & _
					"And ((Round(Dm_Resultsar.Sar*2,4) <> Round(Dm_Padron_Banamex.Salary,4)) " & _
					"Or (Round(Dm_Resultsar.Fovi*2,4) <> Round(Dm_Padron_Banamex.salary_v,4)) " & _
					"Or (Round((Dm_Resultsar.Tot01 + Dm_Resultsar.Tot06),2) <> Round(Dm_Padron_Banamex.Fullpay,2)))" & _ 
					"And (Dm_Padron_Banamex.Status = 1))"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

    If iConnectionType <> ORACLE Then
	sQuery = "Update Dm_Padron_Banamex " & _
				"Set salary = Dm_Resultsar.sar, salary_v=Dm_Resultsar.Fovi, Fullpay = (Dm_Resultsar.Tot01 + Dm_Resultsar.Tot06), " & _
				"Status = 5 " & _
				"From Dm_Resultsar " & _
				"Where (Dm_Padron_banamex.EmployeeID = Dm_Resultsar.EmployeeID) " & _
					"And (Dm_Resultsar.Status = 'c')"
    Else
    sQuery = "Update Dm_Padron_Banamex " & _
				"Set (salary, salary_v, Fullpay, Status)=( " & _
				" Select Dm_Resultsar.sar, Dm_Resultsar.Fovi, Dm_Resultsar.Tot01 + Dm_Resultsar.Tot06, 5 From Dm_Resultsar " & _
				" Where Dm_Padron_banamex.EmployeeID = Dm_Resultsar.EmployeeID " & _
					"And Dm_Resultsar.Status = 'c')"
    End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	sQuery = "Insert Into Dm_Update_Padron_Banamex Select * From Dm_Padron_Banamex Where Status = 5"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	sQuery = "Select * From Dm_Padron_Banamex Where Status=2"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	oRecordset.Close
	Set oRecordset = Nothing
	'Response.Redirect "Reports.asp?ReportID=1035"
	ComparePayrolls = lErrorNumber
	Err.Clear
End Function

Function StartPeriod(oRequest, oADODBConnection, lPeriod, sErrorDescription)
'************************************************************
'Purpose: To start the new period
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "StartPeriod"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim StartPayroll
	Dim EndPayroll
	Dim lNextID
	
	StartPayroll = 0
	EndPayroll = 0
	
	lErrorNumber = GetNewIDFromTable(oADODBConnection, "dm_sar_periods", "PeriodID", "", 1, lNextID, sErrorDescription)	
	lPeriod = oRequest("YearID").Item & "0" & oRequest("PeriodID").Item
	
	sQuery = "Select Min(PayrollID) StartPayroll, Max(PayrollID) EndPayroll From PayrollsCLCs Where (PayrollCode = '" & lPeriod & "')"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If Not IsNull(oRecordset.Fields("StartPayroll").Value) And Not IsNull(oRecordset.Fields("EndPayroll").Value) Then
		StartPayroll = oRecordset.Fields("StartPayroll").Value
		EndPayroll = oRecordset.Fields("EndPayroll").Value
	Else
		lErrorNumber = -1
		sErrorDescription = "No se han cargado CLCs para el periodo señalado"
	End If
	
	If lErrorNumber = 0 Then
		sQuery = "Select Count(*) Total From Dm_Sar_Periods Where (IsOpen = 1)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If oRecordset.Fields("Total") = 0 Then
			sQuery = "Select Count(*) regs From dm_sar_periods Where (periodname = " & lPeriod & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If CLng(oRecordset.Fields("regs").Value) = 0 Then
				sQuery = "Insert Into dm_sar_periods Values(" & lNextID & "," & lPeriod & "," & oRequest("YearID").Item & "," & StartPayroll & "," & EndPayroll & ",1)"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			Else
				sQuery = "Select PeriodName, IsOpen From dm_sar_periods Where periodname = " & lPeriod
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If CInt(oRecordset.Fields("IsOpen").Value) = 1 Then 
					lErrorNumber = -1
					sErrorDescription = "El periodo que trata de abrir ya existe y está abierto"
				Else
					lErrorNumber = -1
					sErrorDescription = "El periodo que trata de abrir ya existe y está cerrado"
				End If
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "Existe un periodo abierto, la carga no puede continuar"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	StartPeriod = lErrorNumber
	Err.Clear
End Function



Function ClosePeriod(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To close the current period
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ClosePeriod"
	Dim asTables
	Dim iIndex
	Dim jIndex
	Dim lNumCols
	Dim lErrorNumber
	Dim lPeriod
	Dim lReportID
	Dim oRecordset
	Dim sDate
	Dim sDocumentName
	Dim sFilePath
	Dim sFileName
	Dim sQuery
	Dim sRowContents
	Dim sTables

	sTables = "Dm_Resultsar,Dm_Resultsar_7s,Dm_Resultsar_Bim01,Dm_Resultsar_Bim02,Dm_Resultsar_Bim03," & _
			  "Dm_Resultsar_Bim04,Dm_Resultsar_Bim05,Dm_Resultsar_Bim06,Dm_Padron_Banamex,Dm_Hist_Nomsar," & _
			  "Dm_Hist_Binmar,Dm_Estr_Qna,Dm_Deleted_HistoryList,Dm_Update_Padron_Banamex"
	asTables = Split(sTables,",")

	If lErrorNumber = 0 Then
		sQuery = "Insert Into Dm_Deleted_HistoryList " & _
				"Select u_version,rfc,curp,SocialSecurityNumber,EmployeeLastName,EmployeeLastName2,EmployeeName,0,CT,birthDate,BirthState, " & _
					"GenderShortName,MaritalStatusID,Address,Colony,City,ZipZone,State,Nombram,EmployeeID,ICEFA,Afore,JoinDate,StartDate, " & _
					"Fovi,WorkingDays,InabilityDays,AbsenceDays,FullPay,Salary,Salary_v,EmployeeContributions,EmployeeContributionsAmount," & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & _
				" From Dm_Padron_Banamex " & _
				"Where (Status = 3)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sQuery = "Delete From Dm_Padron_Banamex Where Status = 3"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_BkUpSAR_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_BkUpSAR_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()
	
	For iIndex = 0 To UBound(asTables)
		sQuery = "Select * From " & asTables(iIndex)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If Not oRecordset.EOF Then
			sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_BkUp_" & asTables(iIndex) & ".txt"
			Do While Not oRecordset.EOF
				lNumCols = oRecordset.Fields.Count
				sRowContents = ""
				For jIndex = 0 To lNumCols - 1
					sRowContents = sRowContents & oRecordset.Fields(jIndex) & TABLE_SEPARATOR
				Next 
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oRecordset.MoveNext
			Loop
		End If
	Next

	If lErrorNumber = 0 Then
		sQuery = "Update Dm_Padron_Banamex Set " & _
				"u_version=Dm_Padron_Banamex_Nuevo.u_version, " & _
				"RFC=Dm_Padron_Banamex_Nuevo.RFC, " & _
				"CURP=Dm_Padron_Banamex_Nuevo.CURP, " & _
				"SocialSecurityNumber=Dm_Padron_Banamex_Nuevo.SocialSecurityNumber, " & _
				"EmployeeLastName=Dm_Padron_Banamex_Nuevo.EmployeeLastName, " & _
				"EmployeeLasName2=Dm_Padron_Banamex_Nuevo.EmployeeLastName2, " & _
				"EmployeeName=Dm_Padron_Banamex_Nuevo.EmployeeName, " & _
				"CT=Dm_Padron_Banamex_Nuevo.CT, " & _
				"BirthDate=Dm_Padron_Banamex_Nuevo.BirthDate, " & _
				"BirthState=Dm_Padron_Banamex_Nuevo.BirthState, " & _
				"GenderShortName=Dm_Padron_Banamex_Nuevo.GenderShortName, " & _
				"JoinDate=Dm_Padron_Banamex_Nuevo.JoinDate, " & _
				"CotDate=Dm_Padron_Banamex_Nuevo.StartDate, " & _
				"Salary=Dm_Padron_Banamex_Nuevo.Salary, " & _
				"Fovi=Dm_Padron_Banamex_Nuevo.Fovi, " & _
				"MaritalStatusID=Dm_Padron_Banamex_Nuevo.MaritalStatusID, " & _
				"Address=Dm_Padron_Banamex_Nuevo.Address, " & _
				"Colony=Dm_Padron_Banamex_Nuevo.Colony, " & _
				"City=Dm_Padron_Banamex_Nuevo.City, " & _
				"ZipZone=Dm_Padron_Banamex_Nuevo.ZipZone, " & _
				"State=Dm_Padron_Banamex_Nuevo.State, " & _
				"Nombram=Dm_Padron_Banamex_Nuevo.Nombram, " & _
				"Afore=Dm_Padron_Banamex_Nuevo.Afore, " & _
				"Salary_v=Dm_Padron_Banamex_Nuevo.Salary_v, " & _
				"FullPay=Dm_Padron_Banamex_Nuevo.FullPay, " & _
				"WorkingDays=Dm_Padron_Banamex_Nuevo.WorkingDays, " & _
				"InabilityDays=Dm_Padron_Banamex_Nuevo.InabilityDays, " & _
				"AbsenceDays=Dm_Padron_Banamex_Nuevo.AbsenceDays, " & _
				"EmployeeContributions=Dm_Padron_Banamex_Nuevo.EmployeeContributions, " & _
				"EmployeeContributionsAmount=Dm_Padron_Banamex_Nuevo.EmployeeContributionsAmount " & _
			"From Dm_Padron_Banamex_Nuevo Where (Dm_Padron_Banamex.EmployeeID = Dm_Padron_Banamex_Nuevo.EmployeeID)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sQuery = "Update Dm_Padron_Banamex Set Status = 1, ChangeFlag = '', UserID = " & aLoginComponent(N_USER_ID_LOGIN)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sQuery = "Delete From Dm_Padron_Banamex Where (EmployeeID Not In (Select EmployeeID From Dm_Padron_Banamex_Nuevo))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sQuery = "Truncate Table Dm_Padron_Banamex_Nuevo"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sQuery = "Select PeriodName From Dm_Sar_Periods Where IsOpen = 1"
		sErrorDescription = "No se pudo generar el respaldo del periodo"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lPeriod = oRecordset.Fields("PeriodName").Value
	End If
	If lErrorNumber = 0 Then
		sQuery = "Truncate Table Dm_Aport_Sar"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sQuery = "Truncate Table Dm_Consar_Info"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sQuery = "Truncate Table DM_Hist_Binmar"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table dm_estr_qna"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim01"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim02"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim03"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim04"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim05"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Resultsar_Bim06"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery =  "Truncate Table Dm_Update_Padron_Banamex"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription,  Null)
	End If
	If lErrorNumber = 0 Then
		sQuery = "Update Dm_Sar_Periods Set IsOpen = 0 Where IsOpen = 1"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If lErrorNumber = 0 Then
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
		sErrorDescription = "No se pudieron guardar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, " & sDate & ", '" & CATALOG_SEPARATOR & "', '', '', '')", "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
	End If
	oEndDate = Now()
	If (lErrorNumber = 0) And B_USE_SMTP Then
		If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	ClosePeriod = lErrorNumber
	Err.Clear
End Function

Function DeletePayrollResume(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To Delete the data of the last resume loaded
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DeletePayrollResume"
	Dim sQuery
	Dim lErrorNumber
	Dim oRecordset
	Dim lPeriod

	sQuery = "Select PeriodName From Dm_Sar_Periods Where (IsOpen = 1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		lPeriod = oRecordset.Fields("PeriodName").Value
		sQuery = "Select Count(*) Total From Dm_Hist_Nomsar Where (PeriodID = " & lPeriod & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If oRecordset.Fields("Total").Value <> 0 Then
			sQuery = "Delete From Dm_Hist_Nomsar Where PeriodID = " & lPeriod
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sQuery = "Delete From Dm_Hist_Binmar Where PeriodID = " & lPeriod
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			lErrorNumber = -1
			sErrorDescription = "No se encontraron datos para el period " & lPeriod
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se encontró ningún bimestre abierto"
	End If
	sQuery = "Delete From Dm_Sar_Periods Where (StartPayroll= 0) Or (EndPayroll=0)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	DeletePayrollResume = lErrorNumber
	Err.Clear
End Function

Function paymentsCompare(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To comapre de payments on the payrolls, CLCs And
'		  loaded payroll resume
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "paymentsCompare"

	Dim sCondition
	Dim sDate
	Dim sDocumentName
	Dim sEmployeeHeader
	Dim sField
	Dim sFileName
	Dim sFilePath
	Dim sGeneralHeader
	Dim sHeaderContents
	Dim sMaxDate
	Dim sMinDate
	Dim sPayrollDate
	Dim sQuery
	Dim sRowContents
	Dim sTruncate
	Dim lLastCompany
	Dim lCurrentPeriod
	Dim lDeductions
	Dim lErrorNumber
	Dim lForPayrollID
	Dim lPayrollID
	Dim lPerceptions
	Dim lPeriodID
	Dim lTotal
	Dim oRecordset
	Dim oRecordsetSumary
	Dim alPayrolls
	Dim asCLCs
	Dim iIndex
	Dim jIndex
	Dim kIndex

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	'Se buscan periodos abiertos
	sQuery = "Select PeriodName From dm_sar_periods Where IsOpen = 1"
	sErrorDescription = "No se encontraron periodos abiertos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se encontraron periodos abiertos"
		lErrorNumber = -1
		If Not oRecordset.EOF Then
		
			lPeriodID = oRecordset.Fields("PeriodName").Value
			lErrorNumber = 0
			
             sQuery = "Select Dbs.CLC, Xls.CLC, Dbs.SocietyID, Dbs.CompanyId, Dbs.PeriodID, Dbs.PaymentDate, Sum(Dbs.Income) Income, Sum(Dbs.Deductions) Deductions, Sum(Dbs.NetIncome) NetIncome, Sum(Dbs.Cpt_01) Cpt_a01, Sum(Dbs.Cpt_04) Cpt_a04, Sum(Dbs.Cpt_05) Cpt_a05, Sum(Dbs.Cpt_06) Cpt_a06, Sum(Dbs.Cpt_07) Cpt_a07, Sum(Dbs.Cpt_08) Cpt_a08, Sum(Dbs.Cpt_11) Cpt_a11, Sum(Dbs.Cpt_44) Cpt_a44, Sum(Dbs.Cpt_B2) Cpt_aB2, Sum(Dbs.Cpt_7S) Cpt_a7S, Sum(Xls.Cpt_01) Cpt_n01, Sum(Xls.Cpt_04) Cpt_n04, Sum(Xls.Cpt_05) Cpt_n05, Sum(Xls.Cpt_06) Cpt_n06, Sum(Xls.Cpt_07) Cpt_n07, Sum(Xls.Cpt_08) Cpt_n08, Sum(Xls.Cpt_11) Cpt_n11, Sum(Xls.Cpt_44) Cpt_n44, Sum(Xls.Cpt_B2) Cpt_nB2, Sum(Xls.Cpt_7S) Cpt_n7S From Dm_Hist_Nomsar Dbs, Dm_Hist_Nomsar Xls, Companies C Where Dbs.PeriodID = (Select PeriodName From Dm_Sar_Periods Where IsOpen = 1) And (Dbs.CLC Not Like 'x%') And (Xls.CLC Like 'x%') And (Dbs.CLC Like SUBSTRING (Xls.CLC,2,15)) And (Dbs.CompanyID = C.CompanyID) Group By Dbs.CLC, Xls.CLC, Dbs.SocietyID, Dbs.CompanyId, Dbs.PeriodID, Dbs.PaymentDate"
			sErrorDescription = "No se pudo obtener la información de los resumenes de nómina"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se encontraron los resumens de nóminas del period " & lPeriodID
				lErrorNumber = -1
				If Not oRecordset.EOF Then
					lErrorNumber = 0
					asCLCs = oRecordset.GetRows
					sQuery = "Truncate Table Dm_Hist_Binmar"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					For iIndex = 0 To UBound(asCLCs,2)
						sQuery = "Insert Into DM_Hist_Binmar (SocietyID, CompanyID, PeriodID, PaymentDate, Income,Deduction,NetIncome, " & _
								"Cpt_a01, Cpt_a04, Cpt_a05, Cpt_a06, Cpt_a07, Cpt_a08, Cpt_a11, Cpt_a44, Cpt_aB2, Cpt_a7S," & _
								"Cpt_n01, Cpt_n04, Cpt_n05, Cpt_n06, Cpt_n07, Cpt_n08, Cpt_n11, Cpt_n44, Cpt_nB2, Cpt_n7S," & _
								"Val_01, Val_04, Val_05, Val_06, Val_07, Val_08, Val_11, Val_44, Val_B2, Val_7S, UserID, LastUpdateDate) Values ("
								For jIndex = 2 To UBound(asCLCs)
									sQuery = sQuery & "|" & asCLCs(jIndex,iIndex) & "|"
								Next
								For jIndex = 9 To UBound(asCLCs)-10
									If asCLC(jIndex,iIndex) = asCLC(jIndex+10,iIndex) Then
										sQuery = sQuery & "|1|"
									Else
										sQuery = sQuery & "|0|"
									End If
								Next
                                sQuery = sQuery & "|" & aLoginComponent(N_USER_ID_LOGIN) & "||" & Left(GetSerialNumberForDate(""), len("00000000"))
								sQuery = sQuery & ")"
								sQuery = Replace(sQuery, "||", ",")
								sQuery = Replace(sQuery, "|", "")
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	paymentsCompare = lErrorNumber
	Err.Clear
End Function

Function distributePayments(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To distribute de payments on the payrolls into
'		  the apropiate period table
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "distributePayments"
	Dim aiPayrolls
	Dim aiPeriods
	Dim iIndex
	Dim oRecordset
	Dim lCurPeriod
	Dim lYear
	Dim lEndPayroll
	Dim lPeriod
	Dim lStartPayroll
	Dim sQuery
	Dim lMaxSalary
	Dim alMaxSalary
	Dim sTable

	aiPeriods = Split("1,1,2,2,3,3,4,4,5,5,6,6", ",", -1, vbBinaryCompare)
	sQuery = "Select Count(*) Total From Dm_Sar_Periods Where (IsOpen = 1)"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If oRecordset.Fields("Total").Value <> 0 Then
		sQuery = "Select PeriodName, PeriodYear, StartPayroll, EndPayroll From Dm_Sar_Periods Where (IsOpen = 1)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lPeriod = oRecordset.Fields("PeriodName").Value
		lStartPayroll = oRecordset.Fields("StartPayroll").Value
		lEndPayroll = oRecordset.Fields("EndPayroll").Value
		lYear = oRecordset.Fields("PeriodYear").Value

		sQuery = "Select Count(*) Total From Payrolls Where (PayrollID >= " & lStartPayroll & ") And (PayrollID <= " & lEndPayroll & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

        'Truncamos tablas a llenar
        sQuery= "Truncate table DM_ESTR_QNA"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR_BIM01"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR_BIM02"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR_BIM03"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR_BIM04"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
         sQuery= "Truncate table DM_RESULTSAR_BIM05"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
        sQuery= "Truncate table DM_RESULTSAR_BIM06"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
        sQuery= "Truncate table DM_RESULTSAR_NEG"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
        sQuery= "Truncate table DM_RESULTSAR_7S"
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

		If oRecordset.Fields("Total").Value <> 0 Then
			sQuery = "Select PayrollId From Payrolls Where (PayrollID >= " & lStartPayroll & ") And (PayrollID <= " & lEndPayroll & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			aiPayrolls = oRecordset.GetRows()
			For iIndex = 0 To UBound(aiPayrolls,2)
				sQuery = "Insert Into Dm_Estr_Qna Select EmployeeID, " & aiPayrolls(0,iIndex) & ", RecordDate, ConceptID, ConceptAmount, 0, 0 From Payroll_" & aiPayrolls(0,iIndex) & " Where conceptid In(1,4,5,6,7,8,13,47,89,146) Order By EmployeeID, ConceptID"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next 
		Else
			lErrorNumber = -1
			sErrorDescription = "No se encontraron los archivos de las nóminas del bimestre " & lPeriod
		End If
		sQuery = "Select Count(*) Total From Payrolls Where (PayrollID >= " & lStartPayroll & "0) And (PayrollID <= " & lEndPayroll & "0)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If oRecordset.Fields("Total").Value <> 0 Then
			sQuery = "Select PayrollId From Payrolls Where (PayrollID >= " & lStartPayroll & ") And (PayrollID <= " & lEndPayroll & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			aiPayrolls = oRecordset.GetRows()
			For iIndex = 0 To UBound(aiPayrolls,2)
				sQuery = "Insert Into Dm_Estr_Qna Select EmployeeID, " & aiPayrolls(0,iIndex) & ", RecordDate, ConceptID, ConceptAmount, 1 From Payroll_" & aiPayrolls(0,iIndex) & " Where conceptid In(1,4,5,6,7,8,13,47,89,146) Order By EmployeeID, ConceptID"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		Else
			lErrorNumber = -1
			sErrorDescription = "No se encontraron los archivos de las nóminas del bimestre " & lPeriod
		End If

        sQuery = "Update Dm_Estr_Qna Set DMS = (Select CurrencyValue From CurrenciesHistoryList Where (Dm_Estr_Qna.RecordDate = CurrenciesHistoryList.CurrencyDate) And (CurrenciesHistoryList.CurrencyID = 1))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sQuery = "Select Distinct RecordDate From Dm_Estr_Qna Order By 1 Asc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aiPayrolls = oRecordset.GetRows()
        lCurPeriod = CLng(Right(CStr(lPeriod),2))
		For iIndex = 1 To 6
			sQuery = "Insert Into Dm_Resultsar_Bim0" & iIndex & " (EmployeeID, u_version, curp, SAR, Fovi, sarExe, FoviExe, Tot01, Tot06, Status, Period)"
			If iIndex <= lCurPeriod Then
				sQuery = sQuery & " Select Distinct Qna.EmployeeID, E.CompanyID, E.Curp, 0,0,0,0,0,0,1," & lPeriod & " From Employees E, Dm_Estr_Qna Qna Where (Qna.EmployeeId = E.EmployeeID) And (RecordDate > " & lYear & Right("00" & (iIndex*2-1),2) & "00) And (RecordDate < " & lYear & Right("00" & (iIndex*2),2) & "99) And (E.CompanyID Is Not Null)"
			Else
				sQuery = sQuery & " Select Distinct Qna.EmployeeID, E.CompanyID, E.Curp, 0,0,0,0,0,0,1," & lPeriod & " From Employees E, Dm_Estr_Qna Qna Where (Qna.EmployeeId = E.EmployeeID) And (RecordDate > " & (CLng(lYear)-1) & Right("00" & (iIndex*2-1),2) & "00) And (RecordDate < " & (CLng(lYear)-1) & Right("00" & (iIndex*2),2) & "99) And (E.CompanyID Is Not Null)"
			End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next
		
		For iIndex = 1 To 6
            If iConnectionType <> ORACLE Then
			sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set Sar = TmpTable.Sar, SarExe = TmpTable.SarExe From " & _
					"(Select EmployeeID, Sum(Sar) Sar, Sum(SarExe) SarExe From " & _
					"(Select EmployeeID, DMS*150 Sar,sum(ConceptAmount) - DMS*150 SarExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,4,5,6,7,8,13,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
					" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 >= 0)) HashTable " & _
					"Group By EmployeeID " & _
					"Union " & _
					"Select EmployeeID, Sum(Sar) Sar, 0 SarExe From " & _
					"(Select EmployeeID, sum(ConceptAmount) Sar, 0 SarExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,4,5,6,7,8,13,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, RecordDate, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 < 0)) HashTable " & _
					"Group By EmployeeID) TmpTable " & _
					"Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = TmpTable.EmployeeID"
            Else
            sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set (Sar,SarExe) = (Select sar,sarexe From " & _
                     "(Select EmployeeID, Sum(sar), Sum(sarexe) From ( " & _
					"Select EmployeeID, Sum(Sar) Sar, Sum(SarExe) SarExe From " & _
					"(Select EmployeeID, DMS*150 Sar,Sum(ConceptAmount) - DMS*150 SarExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,4,5,6,7,8,13,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
					" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, DMS " & _
					"Having (Sum(ConceptAmount) - DMS*150 >= 0)) HashTable " & _
					"Group By EmployeeID " & _
					"Union " & _
					"Select EmployeeID, -Sum(Sar) Sar, 0 SarExe From " & _
					"(Select EmployeeID, sum(ConceptAmount) Sar, 0 SarExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,4,5,6,7,8,13,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, RecordDate, DMS " & _
					"Having (Sum(ConceptAmount) - DMS*150 < 0)) HashTable " & _
					"Group By EmployeeID) " & _
                    " Group By EmployeeID) TmpTable " & _
					"Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = TmpTable.EmployeeID)"
            End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

            If iConnectionType <> ORACLE Then
            sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set Fovi = TmpTable.Fovi, FoviExe = TmpTable.FoviExe From " & _
					"(Select EmployeeID, Sum(Fovi) Fovi, Sum(FoviExe) FoviExe From " & _
					"(Select EmployeeID, DMS*150 Fovi,sum(ConceptAmount) - DMS*150 FoviExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 >= 0)) HashTable " & _
					"Group By EmployeeID " & _
					"Union " & _
					"Select EmployeeID, Sum(Fovi) Fovi, 0 FoviExe From " & _
					"(Select EmployeeID, sum(ConceptAmount) Fovi, 0 FoviExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, RecordDate, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 < 0)) HashTable " & _
					"Group By EmployeeID) TmpTable " & _
					"Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = TmpTable.EmployeeID"
            Else
            'Para el siguiente proceso utilizaremos la tabla temporal DM_RESULTSAR_BIM0X_TMP01
            sQuery = "Select COUNT(*) TOTAL From all_tables Where table_name='DM_RESULTSAR_BIM0X_TMP01' And owner='SIAP'"
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
            If oRecordset.Fields("TOTAL").Value = 1 Then
               sQuery= "Truncate table DM_RESULTSAR_BIM0X_TMP01"
               lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
            Else
               sQuery= "CREATE GLOBAL TEMPORARY TABLE DM_RESULTSAR_BIM0X_TMP01( EMPLOYEEID INTEGER, FOVI NUMBER, FOVIEXE NUMBER)"
               lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
            End If

             DIM sQuery2
             sQuery2 = "Select EmployeeID, Sum(Fovi) Fovi, Sum(FoviExe) FoviExe From (" & _
                    "Select EmployeeID, Sum(Fovi) Fovi, Sum(FoviExe) FoviExe From " & _
					"(Select EmployeeID, DMS*150 Fovi,sum(ConceptAmount) - DMS*150 FoviExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 >= 0)) HashTable " & _
					"Group By EmployeeID " & _
					"Union " & _
					"Select EmployeeID, -Sum(Fovi) Fovi, 0 FoviExe From " & _
					"(Select EmployeeID, sum(ConceptAmount) Fovi, 0 FoviExe " & _
					"From Dm_Estr_Qna " & _
					"Where ConceptID In (1,47,89) " & _
						" And (RecordDate >= " & lYear & Right("00" & ((iIndex*2)-1),2) & "00) " & _
						" And (RecordDate <= " & lYear & Right("00" & ((iIndex*2)),2) & "99) " & _
					"Group By EmployeeID, RecordDate, DMS " & _
					"Having (sum(ConceptAmount) - DMS*150 < 0)) HashTable " & _
                    "Group By EmployeeID) " & _
					"Group By EmployeeID"
	        sQuery = "Insert into DM_RESULTSAR_BIM0X_TMP01 " & sQuery2
	        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set (Fovi, FoviExe) = (Select Fovi, FoviExe From " & _
					"DM_RESULTSAR_BIM0X_TMP01 TmpTable Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = TmpTable.EmployeeID ) " & _
                    "Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID in (Select EmployeeID from DM_RESULTSAR_BIM0X_TMP01)"
            End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

            If iConnectionType <> ORACLE Then
			sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set Tot01 = ConceptAmount From " & _
						"(Select distinct Dm_Estr_Qna.EmployeeID, Dm_Estr_Qna.RecordDate, ConceptAmount " & _
						"From Dm_Estr_Qna, " & _
						"(Select EmployeeID, Max(RecordDate) RecordDate From Dm_Estr_Qna Where ConceptID = 1 Group By EmployeeID) MaxDates " & _
						"Where (Dm_Estr_Qna.EmployeeId = MaxDates.EmployeeID) " & _
							"And (Dm_Estr_Qna.RecordDate = MaxDates.RecordDate) " & _
							"And (Dm_Estr_Qna.ConceptID = 1)) Amount " & _
						"Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = Amount.EmployeeID"
            Else
            sQuery = "Merge Into Dm_Resultsar_Bim0" & iIndex & _
                     " Using(Select distinct Dm_Estr_Qna.EmployeeID, Dm_Estr_Qna.RecordDate, ConceptAmount " & _ 
                     " From Dm_Estr_Qna, ( Select EmployeeID, Max(RecordDate) RecordDate " & _ 
                     " From Dm_Estr_Qna " & _ 
                     " Where ConceptID = 1 " & _ 
                     " Group By EmployeeID " & _
                     " ) MaxDates " & _ 
                     " Where (Dm_Estr_Qna.EmployeeId = MaxDates.EmployeeID) And " & _ 
                     " (Dm_Estr_Qna.RecordDate = MaxDates.RecordDate) And " & _ 
                     " (Dm_Estr_Qna.ConceptID = 1) " & _
                     " ) Amount " & _
                     " On (Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = Amount.EmployeeID) " & _
                     " When Matched Then " & _
                     " Update Set Tot01 = ConceptAmount"
            End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

            If iConnectionType <> ORACLE Then
			sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set Tot06 = ConceptAmount From " & _
						"(Select distinct Dm_Estr_Qna.EmployeeID, Dm_Estr_Qna.RecordDate, ConceptAmount " & _
						"From Dm_Estr_Qna, " & _
						"(Select EmployeeID, Max(RecordDate) RecordDate From Dm_Estr_Qna Where ConceptID = 6 Group By EmployeeID) MaxDates " & _
						"Where (Dm_Estr_Qna.EmployeeId = MaxDates.EmployeeID) " & _
							"And (Dm_Estr_Qna.RecordDate = MaxDates.RecordDate) " & _
							"And (Dm_Estr_Qna.ConceptID = 6)) Amount " & _
						"Where Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = Amount.EmployeeID"
			Else
            sQuery = "MERGE INTO Dm_Resultsar_Bim0" & iIndex & _
                     " using(Select distinct Dm_Estr_Qna.EmployeeID, Dm_Estr_Qna.RecordDate, ConceptAmount " & _ 
                     " From Dm_Estr_Qna, ( Select EmployeeID, Max(RecordDate) RecordDate " & _ 
                     " From Dm_Estr_Qna " & _ 
                     " Where ConceptID = 6 " & _ 
                     " Group By EmployeeID " & _
                     " ) MaxDates " & _ 
                     " Where (Dm_Estr_Qna.EmployeeId = MaxDates.EmployeeID)  And " & _ 
                     " (Dm_Estr_Qna.RecordDate = MaxDates.RecordDate)  And " & _
                     " (Dm_Estr_Qna.ConceptID = 6) " & _
                     " ) Amount " & _ 
                     " On (Dm_Resultsar_Bim0" & iIndex & ".EmployeeID = Amount.EmployeeID) " & _
                     " When Matched Then " & _
                     " Update Set Tot06 = ConceptAmount"
            End If
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next

		For iIndex = 1 To 6
			sQuery = "Update Dm_Resultsar_Bim0" & iIndex & " Set Curp = 0 Where (Curp Is Null)"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next

		sQuery = "Insert Into Dm_Resultsar Select u_version, EmployeeID, Curp, Sar, Fovi, SarExe, FoviExe, Tot01, Tot06, Status, Period, 0 From Dm_Resultsar_Bim0" & lCurPeriod
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sQuery = "Insert Into Dm_ResultSar_7s Select u_version, Qna.EmployeeID, Curp, ConceptAmount, 0, Bx.Status, " & lPeriod & " From Dm_Estr_Qna Qna, Dm_Padron_Banamex Bx Where (Qna.EmployeeID = Bx.EmployeeID) And (ConceptID = 146) Order By Qna.EmployeeID"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

        If iConnectionType <> ORACLE Then 
		    sQuery = "Update Dm_Padron_Banamex Set EmployeeContributionsAmount = Dm_Resultsar_7s.Amount From Dm_Resultsar_7s Where (Dm_Padron_Banamex.EmployeeID = Dm_Resultsar_7s.EmployeeID)"
        Else
            sQuery = "Merge Into Dm_Padron_Banamex A " & _
                     " Using(" & _
                     " Select EmployeeID, sum(amount) amount " & _
                     " From Dm_Resultsar_7s " & _
                     " Group By EMPLOYEEID " & _
                     " ) B " & _
                     " On (A.EmployeeID = B.EmployeeID) " & _
                     " When Matched Then " & _
                     " Update Set A.EmployeeContributionsAmount = B.Amount"
        End If

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sQuery = "Update Dm_Padron_Banamex Set EmployeeContributionsAmount = 0 Where EmployeeContributionsAmount Is Null"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sQuery = "Update Dm_Padron_Banamex Set EmployeeContributions = 1 Where EmployeeContributionsAmount <> 0"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sQuery = "Update Dm_Padron_Banamex Set EmployeeContributions = 0 Where EmployeeContributionsAmount = 0"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PayrollResumeForSarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	Else
		lErrorNumber = -1
		sErrorDescription = "No hay periodos de revisión abiertos"
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	distributePayments = lErrorNumber
	Err.Clear
End Function

Function FormatingTextTabColumns(ReportLoad, Unformatted_File)
'************************************************************
'Purpose: To format the information received from an external 
'flat or cvs file
'Inputs:  ReportLoad, Unformatted_File
'************************************************************
		Dim fso
		Dim oFile_Input
		Dim oFile_Output
		Dim LineaDeTexto
		Dim LineaFinal
        Dim File_Entrada
		Dim File_Salida
		Dim archivo_entrada
		Dim archivo_salida
		Dim Flag
		Dim Contador
		'--------------------------------'Layout de la informacion a subir para ConsarFile
		Dim CF_cve
		Dim CF_filiacion
		Dim CF_curp
		Dim CF_nss
		Dim CF_apellido_p
		Dim CF_apellido_m
		Dim CF_nombre
		Dim CF_nombram
		Dim CF_cve_icefa
		Dim CF_fecha_ing
		Dim CF_fecha_cot
		Dim CF_cred_fov
		Dim CF_dias_cot
		Dim CF_dias_incap
		Dim CF_dias_ausen
		Dim CF_sal_base
		Dim CF_sal_base_v
		Dim CF_sar
		Dim CF_cv_patron
		Dim CF_cv_trabajador
		Dim CF_fov
		Dim CF_ahorro_trab
		Dim CF_ahorro_depe
		'--------------------------------'Layout de la informacion a subir para SarCensus
		Dim SC_u_version
        Dim SC_filiacion
        Dim SC_curp
        Dim SC_nss
        Dim SC_apellido_p
        Dim SC_apellido_m
        Dim SC_nombre
        Dim SC_id_pag
        Dim SC_ct
        Dim SC_fecha_naci
        Dim SC_edo_naci
        Dim SC_sexo
        Dim SC_edo_civil
        Dim SC_domicilio
        Dim SC_col
        Dim SC_pob_del_mun
        Dim SC_cp
        Dim SC_ent_fed
        Dim SC_nombram
        Dim SC_id_empleado
        Dim SC_cve_icefa
        Dim SC_afore
        Dim SC_fecha_ing
        Dim SC_fecha_cot
        Dim SC_fovi
        Dim SC_d_cotiza
        Dim SC_d_incapa
        Dim SC_d_ausent
        Dim SC_sal_inte
        Dim SC_sal_base
        Dim SC_sal_base_v
        Dim SC_aport_emp
        Dim SC_import_emp

		Flag = False
        File_entrada=Unformatted_File 
        File_Salida=Mid(File_entrada, 1,len(File_entrada)-4) & "_Auxiliar.txt"

		archivo_entrada=Server.MapPath("Uploaded Files") & "\" & File_entrada
		archivo_salida=Server.MapPath("Uploaded Files") & "\" & File_Salida

		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set oFile_Input=fso.OpenTextFile(archivo_entrada,1,False)

		If fso.FileExists(archivo_salida)=True Then
		   fso.DeleteFile archivo_salida, True
		End If
		Set oFile_Output=fso.CreateTextFile(archivo_salida)
		Contador=0
		Do While Not oFile_Input.AtEndOfStream
			LineaDeTexto=oFile_Input.ReadLine
			If Contador= 0 And ReportLoad <> "BanamexCensus" Then
				If InStr(LineaDeTexto, Chr(9) ) > 0 Then 
				   Flag= True
				   Exit Do 
				End If
			End If

            If ReportLoad="ConsarFile" Then
			    'Damos formato a la linea de texto de acuerdo a layout proporcionado por el usuario, para ConsarFile
			    CF_cve           = mid(LineaDeTexto, 1  ,  2) & Chr(9)
			    CF_filiacion     = mid(LineaDeTexto, 3  , 13) & Chr(9)
			    CF_curp          = mid(LineaDeTexto, 16 , 18) & Chr(9)
			    CF_nss           = mid(LineaDeTexto, 34 , 11) & Chr(9)
			    CF_apellido_p    = mid(LineaDeTexto, 45 , 40) & Chr(9)
			    CF_apellido_m    = mid(LineaDeTexto, 85 , 40) & Chr(9)
			    CF_nombre        = mid(LineaDeTexto, 125, 40) & Chr(9)
			    CF_nombram       = mid(LineaDeTexto, 165,  1) & Chr(9)
			    CF_cve_icefa     = mid(LineaDeTexto, 166,  3) & Chr(9)
			    CF_fecha_ing     = mid(LineaDeTexto, 169,  8) & Chr(9)
			    CF_fecha_cot     = mid(LineaDeTexto, 177,  8) & Chr(9)
			    CF_cred_fov      = mid(LineaDeTexto, 185,  1) & Chr(9)
			    CF_dias_cot      = mid(LineaDeTexto, 186,  3) & Chr(9)
			    CF_dias_incap    = mid(LineaDeTexto, 189,  3) & Chr(9)
			    CF_dias_ausen    = mid(LineaDeTexto, 192,  3) & Chr(9)
			    CF_sal_base      = mid(LineaDeTexto, 195,  7) & Chr(9)
			    CF_sal_base_v    = mid(LineaDeTexto, 202,  7) & Chr(9)
			    CF_sar           = mid(LineaDeTexto, 209, 12) & Chr(9)
			    CF_cv_patron     = mid(LineaDeTexto, 221, 12) & Chr(9)
			    CF_cv_trabajador = mid(LineaDeTexto, 233, 12) & Chr(9)
			    CF_fov           = mid(LineaDeTexto, 245, 12) & Chr(9)
			    CF_ahorro_trab   = mid(LineaDeTexto, 257, 12) & Chr(9)
			    CF_ahorro_depe   = mid(LineaDeTexto, 269, 12)

			    'Armamos linea donde los campos van separados por tabuladores
			    LineaFinal= CF_cve & CF_filiacion & CF_curp & CF_nss & CF_apellido_p & CF_apellido_m & CF_nombre & CF_nombram & CF_cve_icefa & CF_fecha_ing
			    LineaFinal= LineaFinal & CF_fecha_cot & CF_cred_fov & CF_dias_cot & CF_dias_incap & CF_dias_ausen & CF_sal_base & CF_sal_base_v & CF_sar & CF_cv_patron & CF_cv_trabajador
			    LineaFinal= LineaFinal & CF_fov & CF_ahorro_trab & CF_ahorro_depe
            End If

            If ReportLoad="SarCensus" Then
			    'Damos formato a la linea de texto de acuerdo a layout proporcionado por el usuario, para SarCensus
                SC_u_version   = mid(LineaDeTexto, 1   ,  1) & Chr(9)
                SC_filiacion   = mid(LineaDeTexto, 2   , 13) & Chr(9)
                SC_curp        = mid(LineaDeTexto, 15  , 18) & Chr(9)
                SC_nss         = mid(LineaDeTexto, 33  , 11) & Chr(9)
                SC_apellido_p  = mid(LineaDeTexto, 44  , 40) & Chr(9)
                SC_apellido_m  = mid(LineaDeTexto, 84  , 40) & Chr(9)
                SC_nombre      = mid(LineaDeTexto, 124 , 40) & Chr(9)
                SC_id_pag      = mid(LineaDeTexto, 164 ,  5) & Chr(9)
                SC_ct          = mid(LineaDeTexto, 169 , 20) & Chr(9)
                SC_fecha_naci  = mid(LineaDeTexto, 189 ,  8) & Chr(9)
                SC_edo_naci    = mid(LineaDeTexto, 197 ,  2) & Chr(9)
                SC_sexo        = mid(LineaDeTexto, 199 ,  1) & Chr(9)
                SC_edo_civil   = mid(LineaDeTexto, 200 ,  1) & Chr(9)
                SC_domicilio   = mid(LineaDeTexto, 201 , 60) & Chr(9)
                SC_col         = mid(LineaDeTexto, 261 , 30) & Chr(9)
                SC_pob_del_mun = mid(LineaDeTexto, 291 , 30) & Chr(9)
                SC_cp          = mid(LineaDeTexto, 321 ,  5) & Chr(9)
                SC_ent_fed     = mid(LineaDeTexto, 326 ,  2) & Chr(9)
                SC_nombram     = mid(LineaDeTexto, 328 ,  1) & Chr(9)
                SC_id_empleado = mid(LineaDeTexto, 329 , 10) & Chr(9)
                SC_cve_icefa   = mid(LineaDeTexto, 339 ,  3) & Chr(9)
                SC_afore       = mid(LineaDeTexto, 342 ,  1) & Chr(9)
                SC_fecha_ing   = mid(LineaDeTexto, 343 ,  8) & Chr(9)
                SC_fecha_cot   = mid(LineaDeTexto, 351 ,  8) & Chr(9)
                SC_fovi        = mid(LineaDeTexto, 359 ,  1) & Chr(9)
                SC_d_cotiza    = mid(LineaDeTexto, 360 ,  3) & Chr(9)
                SC_d_incapa    = mid(LineaDeTexto, 363 ,  3) & Chr(9)
                SC_d_ausent    = mid(LineaDeTexto, 366 ,  3) & Chr(9)
                SC_sal_inte    = mid(LineaDeTexto, 369  , 7) & Chr(9)
                SC_sal_base    = mid(LineaDeTexto, 376  , 7) & Chr(9)
                SC_sal_base_v  = mid(LineaDeTexto, 383  , 7) & Chr(9)
                SC_aport_emp   = mid(LineaDeTexto, 390  , 1) & Chr(9)
                SC_import_emp  = mid(LineaDeTexto, 391  , 7)

			    'Armamos linea donde los campos van separados por tabuladores
			    LineaFinal= SC_u_version & SC_filiacion & SC_curp & SC_nss & SC_apellido_p & SC_apellido_m & SC_nombre & SC_id_pag & SC_ct & SC_fecha_naci
			    LineaFinal= LineaFinal & SC_edo_naci & SC_sexo & SC_edo_civil & SC_domicilio & SC_col & SC_pob_del_mun & SC_cp & SC_ent_fed & SC_nombram & SC_id_empleado
                LineaFinal= LineaFinal & SC_cve_icefa & SC_afore & SC_fecha_ing & SC_fecha_cot & SC_fovi & SC_d_cotiza & SC_d_incapa & SC_d_ausent & SC_sal_inte & SC_sal_base
			    LineaFinal= LineaFinal & SC_sal_base_v & SC_aport_emp & SC_import_emp
            End If
            If StrComp(ReportLoad,"BanamexCensus")=0 Then
               If Right(LineaDeTexto,1)="|" Then LineaDeTexto=Left(LineaDeTexto,len(LineaDeTexto)-1)
               LineaDeTexto=replace(LineaDeTexto,"|",Chr(9))
               LineaFinal=LineaDeTexto
            End If
			'Escribimos la linea de texto en el archivo de paso, pero ya formateada, donde las columnas van separadas por tabuladores
			oFile_Output.writeline(LineaFinal)
			Contador= Contador + 1
		Loop
        oFile_Input.close
        oFile_Output.close
		If Flag= False Then
           fso.DeleteFile archivo_entrada, True
           fso.CopyFile archivo_salida,archivo_entrada, True
           fso.DeleteFile archivo_salida, True

		End If
		Set fso= nothing
  End Function
%>