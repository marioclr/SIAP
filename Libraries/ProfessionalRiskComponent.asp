<%
Const N_BRANCH_ID = 0
Const S_BRANCH_SHORT_NAME = 1
Const S_BRANCH_NAME = 2
Const N_CENTER_TYPE_ID_PROFESSIONAL_RISK = 3
Const S_CENTER_TYPE_SHORT_NAME = 4
Const S_CENTER_TYPE_NAME = 5
Const N_POSITION_ID = 6
Const S_POSITION_SHORT_NAME = 7
Const S_POSITION_NAME = 8
Const N_SERVICE_ID = 9
Const S_SERVICE_SHORT_NAME = 10
Const S_SERVICE_NAME = 11
Const N_RISK_LEVEL =12
Const N_START_DATE = 13
Const N_END_DATE = 14
Const N_ACTIVE_PROFESSIONAL_RISK = 15
Const B_CHECK_FOR_DUPLICATED_PROFESSIONAL_RISK = 16
Const B_IS_DUPLICATED_PROFESSIONAL_RISK = 17
Const B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK = 18

Const N_PROESSIONAL_RISK_COMPONENT_SIZE = 18

Dim aProfessionalRiskComponent()
Redim aProfessionalRiskComponent(N_PROESSIONAL_RISK_COMPONENT_SIZE)

Function InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
'************************************************************
'Purpose: To initialize the empty elements of the 
'		  Professional Risk Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aProfessionalRiskComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeProfessionalRiskComponent"
	Dim iItem
	Redim Preserve aProfessionalRiskComponent(N_PROFESSIONAL_RISK_COMPONENT_SIZE)

	aProfessionalRiskComponent(N_BRANCH_ID) = -1
	aProfessionalRiskComponent(N_CENTER_TYPE_ID) = -1
	aProfessionalRiskComponent(N_POSITION_ID) = -1
	aProfessionalRiskComponent(N_SERVICE_ID) = -1
	aProfessionalRiskComponent(S_BRANCH_NAME) = ""
	aProfessionalRiskComponent(S_CENTER_TYPE_NAME) = ""
	aProfessionalRiskComponent(S_POSITION_NAME) = ""
	aProfessionalRiskComponent(S_SERVICE_NAME) = ""
	
	If IsEmpty(aProfessionalRiskComponent(S_BRANCH_SHORT_NAME)) Then
		If Len(oRequest("BranchID").Item) > 0 Then
			aProfessionalRiskComponent(S_BRANCH_SHORT_NAME) = CLng(oRequest("BranchID").Item)
		Else
			aProfessionalRiskComponent(S_BRANCH_SHORT_NAME) = ""
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(S_CENTER_TYPE_SHORT_NAME)) Then
		If Len(oRequest("CenterTypeShortName").Item) > 0 Then
			aProfessionalRiskComponent(S_CENTER_TYPE_SHORT_NAME) = oRequest("CenterTypeShortName").Item
		Else
			aProfessionalRiskComponent(S_CENTER_TYPE_SHORT_NAME) = ""
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(S_POSITION_SHORT_NAME)) Then
		If Len(oRequest("PositionShortName").Item) > 0 Then
			aProfessionalRiskComponent(S_POSITION_SHORT_NAME) = oRequest("PositionShortName").Item
		Else
			aProfessionalRiskComponent(S_POSITION_SHORT_NAME) = ""
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(S_SERVICE_SHORT_NAME)) Then
		If Len(oRequest("ServiceShortName").Item) > 0 Then
			aProfessionalRiskComponent(S_SERVICE_SHORT_NAME) = oRequest("ServiceShortName").Item
		Else
			aProfessionalRiskComponent(S_SERVICE_SHORT_NAME) = ""
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(N_RISK_LEVEL)) Then
		If Len(oRequest("RiskLevel").Item) > 0 Then
			aProfessionalRiskComponent(N_RISK_LEVEL) = oRequest("RiskLevel").Item
		Else
			aProfessionalRiskComponent(N_RISK_LEVEL) = 0
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(N_START_DATE)) Then
		If Len(oRequest("StartDate").Item) > 0 Then
			aProfessionalRiskComponent(N_START_DATE) = oRequest("StartDate").Item
		Else
			aProfessionalRiskComponent(N_START_DATE) = 0
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(N_END_DATE)) Then
		If Len(oRequest("EndDate").Item) > 0 Then
			aProfessionalRiskComponent(N_END_DATE) = oRequest("EndDate").Item
		Else
			aProfessionalRiskComponent(N_END_DATE) = 0
		End If
	End If

	If IsEmpty(aProfessionalRiskComponent(N_ACTIVE)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aProfessionalRiskComponent(N_ACTIVE) = oRequest("Active").Item
		Else
			aProfessionalRiskComponent(N_ACTIVE) = 0
		End If
	End If

	aProfessionalRiskComponent(B_CHECK_FOR_DUPLICATED_PROFESSIONAL_RISK) = True
	aProfessionalRiskComponent(B_IS_DUPLICATED_PROFESSIONAL_RISK) = False

	aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK) = True
	InitializeProfessionalRiskComponent = Err.number
	Err.Clear
End Function

Function AddProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new record into the matrix
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddProfessionalRisk"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
	End If

	If lErrorNumber = 0 Then
		If aProfessionalRiskComponent(B_CHECK_FOR_DUPLICATED_PROFILE) Then
			lErrorNumber = CheckExistencyOfProfessionalRisk(aProfessionalRiskComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aProfessionalRiskComponent(B_IS_DUPLICATED_PROFILE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro para el puesto " & aProfessionalRiskComponent(S_POSITION_SHORT_NAME) & " en el servicio " & aProfessionalRiskComponent(S_SERVICE_SHORT_NAME) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckProfessionalRiskInformationConsistency(aProfessionalRiskComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					lErrorNumber GetProfessionalRisk(oRequest,oADODBConnection,aProfessionalRiskComponent,sErrorDescription)
					sErrorDescription = "No se pudo guardar la información del nuevo registro."
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ProfessionalRiskMatrix (BranchID, PaymentCenterID, PositionID, ServiceID, RiskLevel, StartDate, EndDate, ModifyDate, UserID, Active) Values "(" & aProfessionalRiskComponent(N_BRANCH_ID) & "," & aProfessionalRiskComponent(N_PAYMENT_CENTER_ID) & "," & aProfessionalRiskComponent(N_POSITION_ID) & "," & aProfessionalRiskComponent(N_SERVICE_ID) & "," & aProfessionalRiskComponent(N_RISK_LEVEL) & "," & aProfessionalRiskComponent(N_START_DATE) & "," & aProfessionalRiskComponent(N_END_DATE) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & "," & aProfessionalRiskComponent(N_ACTIVE) & ")", "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	AddProfessionalRisk = lErrorNumber
	Err.Clear
End Function

Function GetProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a record from the matrix
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProfessionalRisk"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
	End If

	If aProfessionalRiskComponent(N_BRANCH_ID) = -1 Or aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) = -1 Or aProfessionalRiskComponent(N_POSITION_ID) = -1 Or aProfessionalRiskComponent(N_SERVICE_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "La información proporcionada no permite ubicar un registro en la matriz de riesgos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ProfessionalRiskMatrix Where (BranchID=" & aProfessionalRiskComponent(N_BRANCH_ID) & ") And (CenterTypeID=" & aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) & ") And (PositionID=" & aProfessionalRiskComponent(N_POSITION_ID) & ") And (ServiceID = " & aProfessionalRiskComponent(N_SERVICE_ID) & ") And (Active = 1)" , "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aProfessionalRiskComponent(N_BRANCH_ID) = oRecordset.Fields("BranchID").Value
				aProfessionalRiskComponent(N_POSITION_ID) = oRecordset.Fields("PositionID").Value
				aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) = oRecordset.Fields("CenterTypeID").Value
				aProfessionalRiskComponent(N_SERVICE_ID) = oRecordset.Fields("ServiceID").Value
				aProfessionalRiskComponent(N_RISK_LEVEL) = oRecordset.Fields("RiskLevel").Value
				aProfessionalRiskComponent(N_START_DATE) = oRecordset.Fields("StartDate").Value
				aProfessionalRiskComponent(N_END_DATE) = oRecordset.Fields("EndDate").Value
			End If
			oRecordset.Close
			If aProfessionalRiskComponent(N_BRANCH_ID) <> -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BranchShortName, BranchName From Branches Where (BranchID=" & aProfessionalRiskComponent(N_BRANCH_ID) & ")" , "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				aProfessionalRiskComponent(S_BRANCH_SHORT_NAME) = oRecordset.Fields("BranchShortName").Value
				aProfessionalRiskComponent(S_BRANCH_NAME) = oRecordset.Fields("BranchName").Value
				oRecordset.Close
			End If
			If aProfessionalRiskComponent(N_POSITION_ID) <> -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionShortName, PositionName From Positions Where (PositionID=" & aProfessionalRiskComponent(N_POSITION_ID) & ")" , "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				aProfessionalRiskComponent(S_POSITION_SHORT_NAME) = oRecordset.Fields("PositionShortName").Value
				aProfessionalRiskComponent(S_POSITION_NAME) = oRecordset.Fields("PositionName").Value
				oRecordset.Close
			End If
			If aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) <> -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CenterTypeShortName, CenterTypeName From CenterTypes Where (CenterTypeID=" & aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) & ")" , "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				aProfessionalRiskComponent(S_CENTER_TYPE_SHORT_NAME) = oRecordset.Fields("CenterTypeShortName").Value
				aProfessionalRiskComponent(S_CENTER_TYPE_NAME) = oRecordset.Fields("CenterTypeName").Value
				oRecordset.Close
			End If
			If aProfessionalRiskComponent(N_SERVICE_ID) <> -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceShortName, ServiceName From Services Where (ServiceID=" & aProfessionalRiskComponent(N_SERVICE_ID) & ")" , "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				aProfessionalRiskComponent(S_SERVICE_SHORT_NAME) = oRecordset.Fields("ServiceShortName").Value
				aProfessionalRiskComponent(S_SERVICE_NAME) = oRecordset.Fields("ServiceName").Value
				oRecordset.Close
			End If
		End If
	End If
	Set oRecordset = Nothing
	GetProfessionalRisk = lErrorNumber
	Err.Clear
End Function

Function GetProfessionalRisks(oRequest, oADODBConnection, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about the proessional risk matrix
'Inputs:  oRequest, oADODBConnection
'Outputs: oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProfessionalRisks"
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ProfessionalRiskMatrix.BranchID,ProfessionalRiskMatrix.CenterTypeID,ProfessionalRiskMatrix.PositionID,ProfessionalRiskMatrix.ServiceID,BranchName,CenterTypeName,PositionShortName,PositionName,ServiceShortName,ServiceName,RiskLevel,ProfessionalRiskMatrix.StartDate,ProfessionalRiskMatrix.EndDate From ProfessionalRiskMatrix, Branches, CenterTypes, Positions, Services Where (ProfessionalRiskMatrix.BranchID = Branches.BranchID) And (ProfessionalRiskMatrix.CenterTypeID=CenterTypes.CenterTypeID) And (ProfessionalRiskMatrix.PositionID=Positions.PositionID) And (ProfessionalRiskMatrix.ServiceID=Services.ServiceID) And (ProfessionalRiskMatrix.Active=1) Order By ProfessionalRiskMatrix.BranchID,ProfessionalRiskMatrix.CenterTypeID,ProfessionalRiskMatrix.PositionID,ProfessionalRiskMatrix.ServiceID ", "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetProfessionalRisks = lErrorNumber
	Err.Clear
End Function

Function ModifyProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing record in the matrix
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyProfessionalRisk"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
	End If

	If aProfessionalRiskComponent(B_CHECK_FOR_DUPLICATED_PROFESSIONAL_RISK) Then
		lErrorNumber = CheckExistencyOfProfessionalRisk(aProfessionalRiskComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If Not CheckProfileInformationConsistency(aProfessionalRiskComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ProfessionalRiskMatrix Set RiskLevel=" & CInt(oRequest("RiskLevel").Item) & " Where (BranchID = " & aProfessionalRiskComponent(N_BRANCH_ID) & ") And (CenterTypeID=" & aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) & ") And (PositionID=" & aProfessionalRiskComponent(N_POSITION_ID) & ") And (ServiceID=" & aProfessionalRiskComponent(N_SERVICE_ID) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyProfessionalRisk = lErrorNumber
	Err.Clear
End Function

Function RemoveProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a record from the matrix
'Inputs:  oRequest, oADODBConnection
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveProfessionalRisk"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
	End If

	If aProfessionalRiskComponent(B_CHECK_FOR_DUPLICATED_PROFESSIONAL_RISK) Then
		lErrorNumber = CheckExistencyOfProfessionalRisk(aProfessionalRiskComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If Not CheckProfessionalRiskInformationConsistency(aProfessionalRiskComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ProfessionalRiskMatrix Where (BranchID = " & aProfessionalRiskComponent(N_BRANCH_ID) & ") And (CenterTypeID=" & aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) & ") And (PositionID=" & aProfessionalRiskComponent(N_POSITION_ID) & ") And (ServiceID=" & aProfessionalRiskComponent(N_SERVICE_ID) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyProfessionalRisk = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfProfessionalRisk(aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record exists in the matrix
'Inputs:  aProfessionalRiskComponent
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfProfessionalRisk"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeProfessionalRiskComponent(oRequest, aProfessionalRiskComponent)
	End If

	If aProfessionalRiskComponent(N_BRANCH_ID) = -1 Or aProfessionalRiskComponent(N_CENTER_TYPE_ID) Or aProfessionalRiskComponent(N_POSITION_ID) Or aProfessionalRiskComponent(N_SERVICE_ID) Then
		lErrorNumber = -1
		sErrorDescription = "La información proporcionada no permite ubicar un registro en la matriz de riesgos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfessionalRiskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ProfessionalRiskLevel Where (BranchID=" & aProfessionalRiskComponent(N_BRANCH_ID) & ") And (CenterTypeID=" & aProfessionalRiskComponent(N_CENTER_TYPE_ID) & ") And (PositionID=" & aProfessionalRiskComponent(N_POSITION_ID) & ") And (ServiceID=" & aProfessionalRiskComponent(N_SERVICE_ID) & ")", "ProfileComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aProfessionalRiskComponent(B_IS_DUPLICATED_PROFILE) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfProfessionalRisk = lErrorNumber
	Err.Clear
End Function

Function CheckProfessionalRiskInformationConsistency(aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the matrix
'Inputs:  aProfessionalRiskComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckProfileInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aProfessionalRiskComponent(N_RISK_LEVEL)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nivel de riesgo no es un valor numérico."
		bIsCorrect = False
	End If
	
	If Not IsNumeric(aProfessionalRiskComponent(N_START_DATE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha inicial no es un valor numérico."
		bIsCorrect = False
	End If

	If Not IsNumeric(aProfessionalRiskComponent(N_END_DATE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha final no es un valor numérico."
		bIsCorrect = False
	End If

	If Len(aProfessionalRiskComponent(S_BRANCH_SHORT_NAME)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha establecido la rama o grupo."
		bIsCorrect = False
	End If
	
	If Len(aProfessionalRiskComponent(S_CENTER_TYPE_SHORT_NAME)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha indicado la unidad administrativa o centro de trabajo."
		bIsCorrect = False
	End If

	If Len(aProfessionalRiskComponent(S_POSITION_SHORT_NAME)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha indicado el puesto."
		bIsCorrect = False
	End If

	If Len(aProfessionalRiskComponent(S_SERVICE_SHORT_NAME)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha indicado el servicio."
		bIsCorrect = False
	End If

	If Len(aProfessionalRiskComponent(N_START_DATE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha establecido la fecha inicial de la vigencia."
		bIsCorrect = False
	End If

	If Len(aProfessionalRiskComponent(N_END_DATE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se ha establecido la fecha fin de vigencia."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos:" & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ProfileComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckProfessionalRiskInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayProfessoionalRiskForm(oRequest, oADODBConnection, sAction, aProfessionalRiskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record from the
'		  matrix using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aProfessionalRiskComponent
'Outputs: aProfessionalRiskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfessoionalRiskForm"
	Dim sNames
	Dim sTempNames
	Dim lErrorNumber
	Dim sPosition
	Dim sService

	If ((Len(oRequest("Modify").Item) <> 0) Or (Len(oRequest("Delete").Item) <> 0)) Then
		If aProfessionalRiskComponent(N_ID_PROFILE) <> -1 Then
			If Len(oRequest("BranchID").Item) > 0 Then 
				aProfessionalRiskComponent(N_BRANCH_ID) = oRequest("BranchID").Item
				aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) = oRequest("CenterTypeID").Item
				aProfessionalRiskComponent(N_POSITION_ID) = oRequest("PositionID").Item
				aProfessionalRiskComponent(N_SERVICE_ID) = oRequest("ServiceID").Item
				aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK) = True
				lErrorNumber = GetProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
			End If
		End If
		If lErrorNumber = 0 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
			Response.Write "<FORM NAME=""ProfessionalRiskFrm"" ID=""ProfessionalRiskFrm"" ACTION=""Catalogs.asp"" METHOD=""POST"" >"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ProfessionalRiskMatrix"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BranchID"" ID=""BranchIDHdn"" VALUE=""" & aProfessionalRiskComponent(N_BRANCH_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterTypeID"" ID=""CenterTypeIDHdn"" VALUE=""" & aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aProfessionalRiskComponent(N_POSITION_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aProfessionalRiskComponent(N_SERVICE_ID) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RiskAmount"" ID=""RiskAmountHdn"" VALUE=""" & aProfessionalRiskComponent(N_RISK_LEVEL) & """ />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Rama:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aProfessionalRiskComponent(S_BRANCH_NAME) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de centro de pago:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aProfessionalRiskComponent(S_CENTER_TYPE_NAME) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						If Len(aProfessionalRiskComponent(S_POSITION_SHORT_NAME)) > 0 Then
							sPosition = "(" & aProfessionalRiskComponent(S_POSITION_SHORT_NAME) & ") " & aProfessionalRiskComponent(S_POSITION_NAME)
						Else
							sPosition = " "
						End If
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & sPosition & " </FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						If Len(aProfessionalRiskComponent(S_SERVICE_SHORT_NAME)) > 0 Then
							sService = "(" & aProfessionalRiskComponent(S_SERVICE_SHORT_NAME) & ") " & aProfessionalRiskComponent(S_SERVICE_NAME)
						Else
							sService = " "
						End If
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & sService & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD COLSPAN=""2""><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						If Len(oRequest("Modify").Item) > 0 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><SELECT NAME=""RiskLevel"" ID=""RiskLevelCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "RiskLevels", "RiskLevelID", "RiskLevelName", "", "RiskLevelName", aProfessionalRiskComponent(N_RISK_LEVEL), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></FONT></TD>"							
							'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""RiskLevel"" ID=""RiskLevelTxt"" VALUE=""" & aProfessionalRiskComponent(N_RISK_LEVEL) & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""TextFields"" /></TD>"
						ElseIf Len(oRequest("Delete").Item > 0) Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aProfessionalRiskComponent(N_RISK_LEVEL) & "</FONT></TD>"
						End If
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR />"

				If aProfessionalRiskComponent(N_ID_PROFILE) = -1 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
				ElseIf Len(oRequest("Delete").Item) > 0 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Remove"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" />"
				Else
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				End If
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Catalogs.asp?Action=ProfessionalRiskMatrix'"" />"
				Response.Write "<BR /><BR />"
				Call DisplayWarningDiv("RemoveCatalogWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
			Response.Write "</FORM>"
		End If
	End If
	DisplayProfessoionalRiskForm = lErrorNumber
	Err.Clear
End Function

Function DisplayProfessionalRiskMatrix(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the matrix in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aProfessionalRiskComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProfessionalRiskMatrix"
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

	iRecordCounter = 1
	lErrorNumber = GetProfessionalRisks(oRequest, oADODBConnection, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Centro de Pago,Rama,Puesto,Servicio,Nivel de Riesgo,Acciones", ",", -1, vbBinaryCompare)
			asCellWidths = Split("20,451,451,451,451,250,50", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
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
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BranchName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CenterTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "(" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & ") " & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "(" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value)) & ") " & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
					If CInt(oRecordset.Fields("RiskLevel").Value) = 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "10"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "20"
					End If
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("BranchID").Value) <> 0 Then
							sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=ProfessionalRiskMatrix&BranchID="&CStr(oRecordset.Fields("BranchID").Value)&"&CenterTypeID="&CStr(oRecordset.Fields("CenterTypeID").Value)&"&PositionID="&oRecordset.Fields("PositionID").Value&"&ServiceID="&CStr(oRecordset.Fields("ServiceID").Value)&"&Modify=1"">"
							'sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=ProfessionalRiskMatrix&BranchID=" & CStr(oRecordset.Fields("BranchID").Value) & "&CenterTypeID=" & CStr(oRecordset.Fields("CenterTypeID").Value) & "&PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & "&ServiceID=" & CStr(oRecordsetFields("ServiceID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_PERMISSIONS) = N_DELETE_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=ProfessionalRiskMatrix&BranchID="&CStr(oRecordset.Fields("BranchID").Value)&"&CenterTypeID="&CStr(oRecordset.Fields("CenterTypeID").Value)&"&PositionID="&oRecordset.Fields("PositionID").Value&"&ServiceID="&CStr(oRecordset.Fields("ServiceID").Value)&"&Delete=1"">"
								'sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=ProfessionalRiskMatrix&BranchID=" & CStr(oRecordset.Fields("BranchID").Value) & "&CenterTypeID=" & CStr(oRecordset.Fields("CenterTypeID").Value) & "&PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & "&ServiceID=" & CStr(oRecordsetFields("ServiceID").Value) & "&Delete=1"">"
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
	DisplayProfessionalRiskMatrix = lErrorNumber
	Err.Clear
End Function
%>