<%
Const N_ID_EMPLOYEE_AUDIT = 0
Const N_ID_CONCEPT_AUDIT = 1
Const N_START_DATE_AUDIT = 2
Const N_AUDIT_ID = 3
Const N_AUDIT_CONCEPT_TYPE_ID = 4
Const N_AUDIT_OPERATION_TYPE = 5
Const N_AUDIT_USER_ID = 6
Const N_AUDIT_DATE = 7
Const S_QUERY_CONDITION_AUDIT = 8
Const B_CHECK_FOR_DUPLICATED_AUDIT = 9
Const B_IS_DUPLICATED_AUDIT = 10
Const B_COMPONENT_INITIALIZED_AUDIT = 11

Const N_AUDIT_COMPONENT_SIZE = 11

Dim aAuditComponent()
ReDim aAuditComponent(N_AUDIT_COMPONENT_SIZE)

Function InitializeAuditComponent(oRequest, aAuditComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Audit Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aAuditComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeAuditComponent"
	Dim oItem
	ReDim Preserve aEmployeeComponent(N_EMPLOYEE_COMPONENT_SIZE)

	If IsEmpty(aAuditComponent(N_ID_EMPLOYEE_AUDIT)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aAuditComponent(N_ID_EMPLOYEE_AUDIT) = CLng(oRequest("EmployeeID").Item)
		Else
			aAuditComponent(N_ID_EMPLOYEE_AUDIT) = -1
		End If
	End If

	If IsEmpty(aAuditComponent(N_ID_CONCEPT_AUDIT)) Then
		If Len(oRequest("ConceptID").Item) > 0 Then
			aAuditComponent(N_ID_CONCEPT_AUDIT) = CLng(oRequest("ConceptID").Item)
		Else
			aAuditComponent(N_ID_CONCEPT_AUDIT) = -1
		End If
	End If
	
	If IsEmpty(aAuditComponent(N_START_DATE_AUDIT)) Then
		If Len(oRequest("ConceptStartYear").Item) > 0 Then
			aAuditComponent(N_START_DATE_AUDIT) = CLng(oRequest("ConceptStartYear").Item & Right(("0" & oRequest("ConceptStartMonth").Item), Len("00")) & Right(("0" & oRequest("ConceptStartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aAuditComponent(N_START_DATE_AUDIT) = CLng(oRequest("ConceptStartDate").Item)
		Else
			aAuditComponent(N_START_DATE_AUDIT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aAuditComponent(N_AUDIT_ID)) Then
		If Len(oRequest("AuditID").Item) > 0 Then
			aAuditComponent(N_AUDIT_ID) = CLng(oRequest("AuditID").Item)
		Else
			aAuditComponent(N_AUDIT_ID) = -1
		End If
	End If

	If IsEmpty(aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID)) Then
		If Len(oRequest("AuditTypeID").Item) > 0 Then
			aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID) = CInt(oRequest("AuditTypeID").Item)
		Else
			aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID) = 3
		End If
	End If

	If IsEmpty(aAuditComponent(N_AUDIT_OPERATION_TYPE)) Then
		If Len(oRequest("AuditOperationTypeID").Item) > 0 Then
			aAuditComponent(N_AUDIT_OPERATION_TYPE) = CInt(oRequest("AuditOperationTypeID").Item)
		Else
			aAuditComponent(N_AUDIT_OPERATION_TYPE) = 3
		End If
	End If

	If IsEmpty(aAuditComponent(N_AUDIT_USER_ID)) Then
		If Len(oRequest("AuditUserID").Item) > 0 Then
			aAuditComponent(N_AUDIT_USER_ID) = CLng(oRequest("AuditUserID").Item)
		Else
			aAuditComponent(N_AUDIT_USER_ID) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aAuditComponent(N_AUDIT_DATE)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aAuditComponent(N_AUDIT_DATE) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aAuditComponent(N_AUDIT_DATE) = CLng(oRequest("StartDate").Item)
		Else
			aAuditComponent(N_AUDIT_DATE) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	aAuditComponent(S_QUERY_CONDITION_AUDIT) = ""
	aAuditComponent(B_CHECK_FOR_DUPLICATED_AUDIT) = False
	aAuditComponent(B_IS_DUPLICATED_AUDIT) = False

	aAuditComponent(B_COMPONENT_INITIALIZED_AUDIT) = True
	InitializeAuditComponent = Err.number
	Err.Clear
End Function

Function AddAudit(oRequest, oADODBConnection, aAuditComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new audit into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAuditComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddAudit"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAuditComponent(B_COMPONENT_INITIALIZED_AUDIT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAuditComponent(oRequest, aAuditComponent)
	End If


	If (aAuditComponent(N_ID_EMPLOYEE_AUDIT) = -1) Or (aAuditComponent(N_ID_CONCEPT_AUDIT) = -1) Or (aAuditComponent(N_START_DATE_AUDIT) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto o la fecha del concepto para registrar la auditoria del movimiento."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AuditComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aAuditComponent(N_AUDIT_ID) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el registro de auditoria."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Audit", "AuditID", "EmployeeID=" & aAuditComponent(N_ID_EMPLOYEE) & " And ConceptID=" & aAuditComponent(N_ID_CONCEPT) & "And StartDate=" & aAuditComponent(N_START_DATE_CONCEPT), 1, aAuditComponent(N_AUDIT_ID), sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			If Not CheckAuditInformationConsistency(aAuditComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar el registro de auditoria."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Audit (EmployeeID, ConceptID, StartDate, AuditID, AuditTypeID, AuditOperationTypeID, AuditUserID, AuditsDate) Values (" & aAuditComponent(N_ID_EMPLOYEE_AUDIT) & ", " & aAuditComponent(N_ID_CONCEPT_AUDIT) & ", " & aAuditComponent(N_START_DATE_AUDIT) & ", " & aAuditComponent(N_AUDIT_ID) & ", " & aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID) & ", " & aAuditComponent(N_AUDIT_OPERATION_TYPE) & ", " & aAuditComponent(N_AUDIT_USER_ID) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "AuditComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddAudit = lErrorNumber
	Err.Clear
End Function

Function CheckAuditInformationConsistency(aAuditComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckAuditInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aAuditComponent(N_ID_EMPLOYEE_AUDIT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del empleado no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAuditComponent(N_ID_CONCEPT_AUDIT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
		bIsCorrect = False
	End If
	
	If Not IsNumeric(aAuditComponent(N_START_DATE_AUDIT)) Then aAuditComponent(N_START_DATE_AUDIT) = Left(GetSerialNumberForDate(""), Len("00000000"))
	
	If Not IsNumeric(aAuditComponent(N_AUDIT_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de auditoria no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del tipo de auditoria no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAuditComponent(N_AUDIT_OPERATION_TYPE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del tipo de operación de auditoria no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAuditComponent(N_AUDIT_USER_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del usuario pra el registro de auditoria no es un valor numérico."
		bIsCorrect = False
	End If

	If Not IsNumeric(aAuditComponent(N_AUDIT_DATE)) Then aAuditComponent(N_AUDIT_DATE) = Left(GetSerialNumberForDate(""), Len("00000000"))

	CheckAuditInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function GetAudit(oRequest, oADODBConnection, aAuditComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a audit from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAuditComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAudit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAuditComponent(B_COMPONENT_INITIALIZED_AUDIT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAuditComponent(oRequest, aAuditComponent)
	End If

	If aAuditComponent(N_ID_EMPLOYEE_AUDIT) = -1 Or aAuditComponent(N_ID_CONCEPT_AUDIT) = -1 Or aAuditComponent(N_START_DATE_AUDIT) = -1 Or aAuditComponent(N_AUDIT_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AuditComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Audit Where EmployeeID=" & aAuditComponent(N_ID_CONCEPT_AUDIT) & " And ConceptID=" & aAuditComponent(N_ID_CONCEPT_AUDIT) & " And StartDate=" & aAuditComponent(N_START_DATE_AUDIT) & " And AuditID=" & aAuditComponent(N_AUDIT_ID), "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AuditComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID) = CLng(oRecordset.Fields("AuditTypeID").Value)
				aAuditComponent(N_AUDIT_OPERATION_TYPE) = CStr(oRecordset.Fields("AuditOperationTypeID").Value)
				aAuditComponent(N_AUDIT_USER_ID) = CStr(oRecordset.Fields("AuditUserID").Value)
				aAuditComponent(N_AUDIT_DATE) = CLng(oRecordset.Fields("AuditsDate").Value)
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetAudit = lErrorNumber
	Err.Clear
End Function

Function GetAudits(oRequest, oADODBConnection, aAuditComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAudits"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAuditComponent(B_COMPONENT_INITIALIZED_AUDIT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAuditComponent(oRequest, aAuditComponent)
	End If

	If (Len(aAuditComponent(S_QUERY_CONDITION_AUDIT)) > 0) Then
		sCondition = Trim(aConceptComponent(S_QUERY_CONDITION_AUDIT))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Audit Where (AuditID>-760211) " & sCondition & " Order By EmployeeID, ConceptID, StartDate, AuditID", "AuditComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetAudits = lErrorNumber
	Err.Clear
End Function
%>
