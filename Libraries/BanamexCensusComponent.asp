<!-- #include file="ZIPLibrary.asp" -->
<%
Const N_U_VERSION = 0
Const N_EMPLOYEE_ID = 1
Const S_RFC = 2
Const S_CURP = 3
Const S_SOCIAL_SECURITY_NUMBER = 4
Const S_EMPLOYEE_LAST_NAME = 5
Const S_EMPLOYEE_LAST_NAME_2 = 6
Const S_EMPLOYEE_NAME = 7
Const S_CT = 8
Const N_BIRTH_DATE = 9
Const N_BIRTH_STATE_ID = 10
Const S_BIRTH_STATE_NAME = 11
Const S_GENDER_SHORT_NAME = 12
Const N_JOIN_DATE = 13
Const N_COT_DATE = 14
Const N_SALARY = 15
Const N_FOVY = 16
Const N_PERIOD_ID = 17
Const N_STATUS_ID = 18
Const N_CHANGE_FLAG = 19
Const N_MARITAL_STATUS_ID = 20
Const S_ADDRESS = 21
Const S_COLONY = 22
Const S_CITY = 23
Const N_ZIP_ZONE = 24
Const S_STATE = 25
Const N_NOMBRAM = 26
Const N_AFORE = 27
Const N_ICEFA = 28
Const N_IC_NUMBER = 29
Const S_MOT_BAJA = 30
Const N_SALARY_V = 31
Const N_FULL_PAY = 32
Const N_WORKING_DAYS = 33
Const N_INABILITY_DAYS = 34
Const N_ABSENCE_DAYS = 35
Const N_EMPLOYEE_CONTRIBUTIONS = 36
Const N_EMPLOYEE_CONTRIBUTIONS_AMOUNT = 37
Const N_START_DATE_FOR_CENSUS = 38
Const N_END_DATE_FOR_CENSUS = 39
Const N_USER_ID_FOR_CENSUS_ID = 40
Const N_LAST_UPDATE_DATE = 41
Const S_COMMENTS_FOR_CENSUS = 42

Const B_CHECK_FOR_DUPLICATED_BANAMEX_CENSUS = 43
Const B_IS_DUPLICATED_BANAMEX_CENSUS = 44
Const B_COMPONENT_INITIALIZED_BANAMEX_CENSUS = 45

Const N_BANAMEX_CENSUS_COMPONENT_SIZE = 45

Dim aBanamexCensusComponent()
Redim aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)

Function InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
'************************************************************
'Purpose: To initialize the empty elements of the 
'		  Banamex Census Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aBanamexCensusComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeBanamexCensusComponent"
	Dim iItem
	Redim Preserve aBanamexCensusComponent(N_BANAMEX_CENSUS_COMPONENT_SIZE)

	aBanamexCensusComponent(N_U_VERSION) = -1
	aBanamexCensusComponent(N_EMPLOYEE_ID) = -1
	aBanamexCensusComponent(S_RFC) = ""
	aBanamexCensusComponent(S_CURP) = ""
	aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) = ""
	aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) = ""
	aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) = ""
	aBanamexCensusComponent(S_EMPLOYEE_NAME) = ""
	aBanamexCensusComponent(S_CT) = ""
	aBanamexCensusComponent(N_BIRTH_DATE) = 0
	aBanamexCensusComponent(N_BIRTH_STATE_ID) = -1
	aBanamexCensusComponent(S_BIRTH_STATE_NAME) = ""
	aBanamexCensusComponent(S_GENDER_SHORT_NAME) = ""
	aBanamexCensusComponent(N_JOIN_DATE) = -1
	aBanamexCensusComponent(N_COT_DATE) = -1
	aBanamexCensusComponent(N_SALARY) = -1
	aBanamexCensusComponent(N_FOVY) = -1
	aBanamexCensusComponent(N_PERIOD_ID) = -1
	aBanamexCensusComponent(N_STATUS_ID) = -1
	aBanamexCensusComponent(N_CHANGE_FLAG) = -1
	aBanamexCensusComponent(N_MARITAL_STATUS_ID) = -1
	aBanamexCensusComponent(S_ADDRESS) = ""
	aBanamexCensusComponent(S_COLONY) = ""
	aBanamexCensusComponent(S_CITY) = ""
	aBanamexCensusComponent(N_ZIP_ZONE) = -1
	aBanamexCensusComponent(S_STATE) = ""
	aBanamexCensusComponent(N_NOMBRAM) = -1
	aBanamexCensusComponent(N_AFORE) = -1
	aBanamexCensusComponent(N_ICEFA) = -1
	aBanamexCensusComponent(N_IC_NUMBER) = -1
	aBanamexCensusComponent(S_MOT_BAJA) = ""
	aBanamexCensusComponent(N_SALARY_V) = -1
	aBanamexCensusComponent(N_FULL_PAY) = -1
	aBanamexCensusComponent(N_WORKING_DAYS) = -1
	aBanamexCensusComponent(N_INABILITY_DAYS) = -1
	aBanamexCensusComponent(N_ABSENCE_DAYS) = -1
	aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) = -1
	aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) = -1
	aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) = -1
	aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) = -1
	aBanamexCensusComponent(N_USER_ID_FOR_CENSUS_ID) = -1
	aBanamexCensusComponent(N_LAST_UPDATE_DATE) = -1
	aBanamexCensusComponent(S_COMMENTS_FOR_CENSUS) = ""

	aBanamexCensusComponent(B_CHECK_FOR_DUPLICATED_BANAMEX_CENSUS) = True
	aBanamexCensusComponent(B_IS_DUPLICATED_BANAMEX_CENSUS) = False

	aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS) = True
	InitializeBanamexCensusComponent = Err.number
	Err.Clear
End Function

Function AddBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new record into the table and sets status
'			field with 2
'Inputs:  oRequest, oADODBConnection
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddBanamexCensusRecord"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If
	If lErrorNumber = 0 Then
		If aBanamexCensusComponent(B_CHECK_FOR_DUPLICATED_BANAMEX_CENSUS) Then
			lErrorNumber = CheckExistencyOfRecord(aBanamexCensusComponent, sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			If aBanamexCensusComponent(B_IS_DUPLICATED_BANAMEX_CENSUS) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro para el empleado " & aBanamexCensusComponent(N_EMPLOYEE_ID)
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BanamexCensusComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckBanamexCensusInformationConsistency(aBanamexCensusComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					aBanamexCensusComponent(N_STATUS_ID) = 2
					sQuery = "Insert Into DM_PADRON_BANAMEX (u_version,EmployeeID,RFC,CURP,SocialSecurityNumber," & _
							"EmployeeLastName,EmployeeLastName2,EmployeeName,CT,BirthDate,BirthState,GenderShortName," & _
							"JoinDate,CotDate,Salary,Fovi,Period,Status,ChangeFlag,MaritalStatusID,Address,Colony," & _
							"City,ZipZone,State,Nombram,Afore,ICEFA,ICNumber,mot_baja,Salary_v,FullPay,WorkingDays," & _
							"InabilityDays,AbsenceDays,EmployeeContributions,EmployeeContributionsAmount,StartDate," & _
							"EndDate,UserID,LastUpdateDate,Comments) Values ('" & _
							aBanamexCensusComponent(N_U_VERSION) & "'," & aBanamexCensusComponent(N_EMPLOYEE_ID) & ",'" & _
							aBanamexCensusComponent(S_RFC) & "','" & aBanamexCensusComponent(S_CURP) & "','" & _
							aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) & "','" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) & "','" & _
							aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) & "','" & aBanamexCensusComponent(S_EMPLOYEE_NAME) & "','" & _
							aBanamexCensusComponent(S_CT) & "'," & aBanamexCensusComponent(N_BIRTH_DATE) & "," & _
							aBanamexCensusComponent(N_BIRTH_STATE_ID) & ",'" & aBanamexCensusComponent(S_GENDER_SHORT_NAME) & "'," & _
							aBanamexCensusComponent(N_JOIN_DATE) & "," & aBanamexCensusComponent(N_COT_DATE) & "," & _
							aBanamexCensusComponent(N_SALARY) & "," & _
							aBanamexCensusComponent(N_FOVY) & "," & aBanamexCensusComponent(N_PERIOD_ID) & "," & _
							aBanamexCensusComponent(N_STATUS_ID) & "," & aBanamexCensusComponent(N_CHANGE_FLAG) & "," & _
							aBanamexCensusComponent(N_MARITAL_STATUS_ID) & ",'" & aBanamexCensusComponent(S_ADDRESS) & "','" & _
							aBanamexCensusComponent(S_COLONY) & "','" & aBanamexCensusComponent(S_CITY) & "'," & _
							aBanamexCensusComponent(N_ZIP_ZONE) & ",'" & aBanamexCensusComponent(S_STATE) & "'," & _
							aBanamexCensusComponent(N_NOMBRAM) & "," & aBanamexCensusComponent(N_AFORE) & "," & _
							aBanamexCensusComponent(N_ICEFA) & "," & aBanamexCensusComponent(N_IC_NUMBER) & ",'" & _
							aBanamexCensusComponent(S_MOT_BAJA) & "'," & aBanamexCensusComponent(N_SALARY_V) & "," & _
							aBanamexCensusComponent(N_FULL_PAY) & "," & aBanamexCensusComponent(N_WORKING_DAYS) & "," & _
							aBanamexCensusComponent(N_INABILITY_DAYS) & "," & aBanamexCensusComponent(N_ABSENCE_DAYS) & "," & _
							aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) & "," & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) & "," & _
							aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) & "," & aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) & "," & _
							aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",'" & _
							Replace(aBanamexCensusComponent(S_COMMENTS_FOR_CENSUS), "'", "") & "')"
					sErrorDescription = "No se pudo guardar la información del nuevo registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If
	AddBanamexCensusRecord = lErrorNumber
	Err.Clear
End Function

Function GetBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a record from the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBanamexCensusRecord"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If

	If aBanamexCensusComponent(N_EMPLOYEE_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se ha indicado el número de empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BanamexCensusComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From DM_PADRON_BANAMEX Where (EmployeeID=" & oRequest("EmployeeID").Item & ")", "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado indicado no se encuentra en la base de datos."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BanamexCensusComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aBanamexCensusComponent(N_U_VERSION) = oRecordset.Fields("u_version").Value
				aBanamexCensusComponent(S_RFC) = oRecordset.Fields("RFC").Value
				aBanamexCensusComponent(S_CURP) = oRecordset.Fields("CURP").Value
				aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) = oRecordset.Fields("SocialSecurityNumber").Value
				aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) = oRecordset.Fields("EmployeeLastName").Value
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) = oRecordset.Fields("EmployeeLastName2").Value
				Else
					aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2)  = ""
				End If
				aBanamexCensusComponent(S_EMPLOYEE_NAME) = oRecordset.Fields("EmployeeName").Value
				If (IsNull(oRecordset.Fields("CT").Value)) Or (oRecordset.Fields("CT").Value = "") Then
					aBanamexCensusComponent(S_CT) = 0
				Else
					aBanamexCensusComponent(S_CT) = oRecordset.Fields("CT").Value
				End If
				aBanamexCensusComponent(N_BIRTH_DATE) = oRecordset.Fields("BirthDate").Value
				aBanamexCensusComponent(N_BIRTH_STATE_ID) = oRecordset.Fields("BirthState").Value
				aBanamexCensusComponent(S_GENDER_SHORT_NAME) = oRecordset.Fields("GenderShortName").Value
				aBanamexCensusComponent(N_JOIN_DATE) = oRecordset.Fields("JoinDate").Value
				aBanamexCensusComponent(N_COT_DATE) = oRecordset.Fields("CotDate").Value
				aBanamexCensusComponent(N_SALARY) = oRecordset.Fields("Salary").Value
				aBanamexCensusComponent(N_FOVY) = oRecordset.Fields("Fovi").Value
				aBanamexCensusComponent(N_PERIOD_ID) = oRecordset.Fields("Period").Value
				aBanamexCensusComponent(N_STATUS_ID) = oRecordset.Fields("Status").Value
				aBanamexCensusComponent(N_CHANGE_FLAG) = oRecordset.Fields("ChangeFlag").Value
				aBanamexCensusComponent(N_MARITAL_STATUS_ID) = oRecordset.Fields("MaritalStatusID").Value
				aBanamexCensusComponent(S_ADDRESS) = oRecordset.Fields("Address").Value
				aBanamexCensusComponent(S_COLONY) = oRecordset.Fields("Colony").Value
				aBanamexCensusComponent(S_CITY) = oRecordset.Fields("City").Value
				aBanamexCensusComponent(N_ZIP_ZONE) = oRecordset.Fields("ZipZone").Value
				aBanamexCensusComponent(S_STATE) = oRecordset.Fields("State").Value
				aBanamexCensusComponent(N_NOMBRAM) = oRecordset.Fields("Nombram").Value
				aBanamexCensusComponent(N_AFORE) = oRecordset.Fields("Afore").Value
				aBanamexCensusComponent(N_ICEFA) = oRecordset.Fields("ICEFA").Value
				aBanamexCensusComponent(N_IC_NUMBER) = oRecordset.Fields("ICNumber").Value
				If (IsNull(oRecordset.Fields("mot_baja").Value)) Or (oRecordset.Fields("mot_baja").Value="") Then
					aBanamexCensusComponent(S_MOT_BAJA) = 0
				Else
					aBanamexCensusComponent(S_MOT_BAJA) = oRecordset.Fields("mot_baja").Value
				End If
				aBanamexCensusComponent(N_SALARY_V) = oRecordset.Fields("Salary_v").Value
				aBanamexCensusComponent(N_FULL_PAY) = oRecordset.Fields("FullPay").Value
				aBanamexCensusComponent(N_WORKING_DAYS) = oRecordset.Fields("WorkingDays").Value
				If (IsNull(oRecordset.Fields("InabilityDays").Value)) Or (oRecordset.Fields("InabilityDays").Value="") Then
					aBanamexCensusComponent(N_INABILITY_DAYS) = 0
				Else
					aBanamexCensusComponent(N_INABILITY_DAYS) = oRecordset.Fields("InabilityDays").Value
				End If
				If (IsNull(oRecordset.Fields("AbsenceDays").Value)) Or (oRecordset.Fields("AbsenceDays").Value="") Then
					aBanamexCensusComponent(N_ABSENCE_DAYS) = 0
				Else
					aBanamexCensusComponent(N_ABSENCE_DAYS) = oRecordset.Fields("AbsenceDays").Value
				End If				
				
				If IsNull(oRecordset.Fields("EmployeeContributions").Value) Or (oRecordset.Fields("EmployeeContributions").Value="") Then
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) = 0
				Else
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) = oRecordset.Fields("EmployeeContributions").Value
				End If
				If IsNull(oRecordset.Fields("EmployeeContributionsAmount").Value) Or (oRecordset.Fields("EmployeeContributionsAmount").Value="") Then
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) = 0
				Else
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) = oRecordset.Fields("EmployeeContributionsAmount").Value
				End If
				aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) = oRecordset.Fields("StartDate").Value
				aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) = oRecordset.Fields("EndDate").Value
				aBanamexCensusComponent(N_USER_ID_FOR_CENSUS_ID) = oRecordset.Fields("UserID").Value
				aBanamexCensusComponent(N_LAST_UPDATE_DATE) = oRecordset.Fields("LastUpdateDate").Value
				aBanamexCensusComponent(S_COMMENTS_FOR_CENSUS) = oRecordset.Fields("Comments").Value
			End If
			oRecordset.Close
		End If
	End If
	Set oRecordset = Nothing
	GetBanamexCensusRecord = lErrorNumber
	Err.Clear
End Function

Function GetBanamexCensusList(oRequest, oADODBConnection, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about the banamex census
'Inputs:  oRequest, oADODBConnection
'Outputs: oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBanamexCensusList"
	Dim lErrorNumber
	Dim sCondition
	Dim sQuery
	sCondition = ""
	If Len(oRequest("EmployeeNumber").Item) > 0 Then 
		If Len(oRequest("EmployeeNumber").Item) = 6 Then
			sCondition = sCondition & " EmployeeID Like '" & CLng(oRequest("EmployeeNumber").Item) & "'|"  
		Else
			sCondition = sCondition & " EmployeeID Like '%" & oRequest("EmployeeNumber").Item & "%'|" 
		End If
	End If
	If Len(oRequest("EmployeeName").Item) > 0 Then sCondition = sCondition & " EmployeeName Like '%" & oRequest("EmployeeName").Item & "%'|"
	If Len(oRequest("EmployeeLastName").Item) > 0 Then sCondition = sCondition & " EmployeeLastName Like '%" & oRequest("EmployeeLastName").Item & "%'|"
	If Len(oRequest("EmployeeLastName2").Item) > 0 Then sCondition = sCondition & " EmployeeLastName2 Like '%" & oRequest("EmployeeLastName2").Item & "%'|"
	If Len(oRequest("RFC").Item) > 0 Then sCondition = sCondition & " RFC Like '%" & oRequest("RFC").Item & "%'|"
	If Len(oRequest("CURP").Item) > 0 Then sCondition = sCondition & " CURP Like '%" & oRequest("CURP").Item & "%'|"
	If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " u_version = " & oRequest("CompanyID").Item & "|"
	If Len(oRequest("Status").Item) > 0 Then sCondition = sCondition & " u_version = " & oRequest("CompanyID").Item & "|"
	
	If Len(sCondiction) > 0 Then
		If StrComp(oRequest("Action").Item, "EmployeesDeleted", vbBinaryCompare) <> 0 Then
			sCondition = " And " & Replace(Mid(sCondition, 1, (Len(sCondition) - 1)), "|", " And ")
		Else
			sCondition = " Where " & Replace(Mid(sCondition, 1, (Len(sCondition) - 1)), "|", " And ")
		End If
	End If

	sErrorDescription = "No se pudo obtener la información de los registros."
	If StrComp(oRequest("Action").Item,"EmployeesDeleted",vbBinaryCompare) = 0 Then
		sQuery = "Select * From DM_DELETED_HISTORYLIST " & sCondition & " Order By EmployeeID"
	Else
		sQuery = "Select * From DM_PADRON_BANAMEX Where (Status<>3) " & sCondition & " Order By EmployeeID"
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetBanamexCensusList = lErrorNumber
	Err.Clear
End Function

Function ModifyBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing record in the table and 
'		mark its status field with 4
'Inputs:  oRequest, oADODBConnection
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyBanamexCensusRecord"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If

	If aBanamexCensusComponent(B_CHECK_FOR_DUPLICATED_BANAMEX_CENSUS) Then
		lErrorNumber = CheckExistencyOfRecord(aBanamexCensusComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If Not CheckBanamexCensusInformationConsistency(aBanamexCensusComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sQuery = "Update dm_padron_banamex Set LastUpdateDate = " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & " Where EmployeeID = " & aBanamexCensusComponent(N_EMPLOYEE_ID) 
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sQuery = "Insert Into dm_update_padron_banamex select * From dm_padron_banamex where employeeID = " & aBanamexCensusComponent(N_EMPLOYEE_ID)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sQuery = "Update DM_PADRON_BANAMEX Set u_version='" & aBanamexCensusComponent(N_U_VERSION) & _
			"',EmployeeID=" & aBanamexCensusComponent(N_EMPLOYEE_ID) & ",RFC='" & aBanamexCensusComponent(S_RFC) & _
			"',CURP='" & aBanamexCensusComponent(S_CURP) & "',SocialSecurityNumber=" & aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) & _
			",EmployeeLastName='" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) & _
			"',EmployeeLastName2='" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) & _
			"',EmployeeName='" & aBanamexCensusComponent(S_EMPLOYEE_NAME) & _
			"',BirthDate=" & aBanamexCensusComponent(N_BIRTH_DATE) & _
			",BirthState=" & aBanamexCensusComponent(N_BIRTH_STATE_ID) & _
			",GenderShortName='" & aBanamexCensusComponent(S_GENDER_SHORT_NAME) & _
			"',JoinDate=" & aBanamexCensusComponent(N_JOIN_DATE) & ",CotDate=" & aBanamexCensusComponent(N_COT_DATE) & _
			",Salary=" & aBanamexCensusComponent(N_SALARY) & ",Fovi=" & aBanamexCensusComponent(N_FOVY) & _
			",Period=" & aBanamexCensusComponent(N_PERIOD_ID) & ",Status=4" & _
			",ChangeFlag='" & aBanamexCensusComponent(N_CHANGE_FLAG) & _
			"',MaritalStatusID=" & aBanamexCensusComponent(N_MARITAL_STATUS_ID) & _
			",Address='" & aBanamexCensusComponent(S_ADDRESS) & "',Colony='" & aBanamexCensusComponent(S_COLONY) & _
			"',City='" & aBanamexCensusComponent(S_CITY) & "',ZipZone=" & aBanamexCensusComponent(N_ZIP_ZONE) & _
			",State='" & aBanamexCensusComponent(S_STATE) & "',Nombram=" & aBanamexCensusComponent(N_NOMBRAM) & _
			",Afore=" & aBanamexCensusComponent(N_AFORE) & ",ICEFA=" & aBanamexCensusComponent(N_ICEFA) & _
			",ICNumber=" & aBanamexCensusComponent(N_IC_NUMBER) & ",mot_baja='" & aBanamexCensusComponent(S_MOT_BAJA) & _
			"',Salary_v=" & aBanamexCensusComponent(N_SALARY_V) & ",FullPay=" & aBanamexCensusComponent(N_FULL_PAY) & _
			",WorkingDays=" & aBanamexCensusComponent(N_WORKING_DAYS) & _
			",InabilityDays=" & aBanamexCensusComponent(N_INABILITY_DAYS) & _
			",AbsenceDays=" & aBanamexCensusComponent(N_ABSENCE_DAYS) & _
			",EmployeeContributions=" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) & _
			",EmployeeContributionsAmount=" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) & _
			",StartDate=" & aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) & ",EndDate=" & aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) & _
			" Where EmployeeID=" & aBanamexCensusComponent(N_EMPLOYEE_ID)

			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyBanamexCensusRecord = lErrorNumber
	Err.Clear
End Function

Function MarkRecordForDeleting(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To mark the status field with 3 for an existing record 
'			in the table for deleting during reports
'Inputs:  oRequest, oADODBConnection
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MarkRecordForDeleting"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo eliminar el registro indicado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update DM_PADRON_BANAMEX Set Status = 3 Where EmployeeID = " & aBanamexCensusComponent(N_EMPLOYEE_ID), "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	MarkRecordForDeleting = lErrorNumber
	Err.Clear
	
End Function

Function RemoveMarkedRecords(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To remove the records marked in the table
'Inputs:  oRequest, oADODBConnection
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveMarkedRecords"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo eliminar el registro indicado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From DM_PADRON_BANAMEX Where Status = 3", "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveMarkedRecords = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfRecord(aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record exists in the table
'Inputs:  aBanamexCensusComponent
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfRecord"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
	End If

	If aBanamexCensusComponent(N_SOCIETY_ID) = -1 Or aBanamexCensusComponent(N_COMPANY_ID) = -1 Or aBanamexCensusComponent(N_BANK_ID) = -1 Or aBanamexCensusComponent(N_PAYMENT_DATE) = -1 Or aBanamexCensusComponent(N_EMPLOYEE_TYPE_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "La información proporcionada no permite ubicar un registro en el padrón."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BanamexCensusComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From DM_PADRON_BANAMEX Where (SocietyID=" & oRequest("SocietyID").Item & ") And (CompanyID=" & oRequest("CompanyID").Item & ") And (PaymentDate=" & oRequest("PaymentDate").Item & ") And (BankID=" & oRequest("BankID").Item & ") And (EmployeeTypeID=" & oRequest("EmployeeTypeID").Item & ")", "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aBanamexCensusComponent(B_IS_DUPLICATED_BANAMEX_CENSUS) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfRecord = lErrorNumber
	Err.Clear
End Function

Function CheckBanamexCensusInformationConsistency(aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the matrix
'Inputs:  aBanamexCensusComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckProfileInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aBanamexCensusComponent(N_EMPLOYEE_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de empleado debe ser numérico."
		bIsCorrect = False
	End If
	
	If Len(Trim(aBanamexCensusComponent(S_RFC))) <> 13 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;-	El RFC proporcionado no está completo."
		bIsCorrect = False
	End If

	If Len(Trim(aBanamexCensusComponent(S_CURP))) <> 18 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;-	La CURP proporcionado no está completo."
		bIsCorrect = False
	End If
	
	If Not IsNumeric(aBanamexCensusComponent(N_SALARY)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del salario no es numérico."
		bIsCorrect = False
	End If
	
	If Not IsNumeric(aBanamexCensusComponent(N_SALARY_V)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor salario_v no es numérico."
		bIsCorrect = False
	End If

	CheckBanamexCensusInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayBanamexCensusForm(oRequest, oADODBConnection, sAction, aBanamexCensusComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record from the
'		  table using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aBanamexCensusComponent
'Outputs: aBanamexCensusComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBanamexCensusForm"
	Dim sNames
	Dim sTempNames
	Dim lCurrentDate
	Dim lErrorNumber
	Dim sPosition
	Dim sService

	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If ((Len(oRequest("Modify").Item) <> 0) Or (Len(oRequest("Delete").Item) <> 0) (Len(oRequest("AddNew").Item) <> 0)) Then
		If aBanamexCensusComponent(N_EMPLOYEE_ID) = -1 Then
			If Len(oRequest("EmployeeID").Item) > 0 Then 
				aBanamexCensusComponent(N_EMPLOYEE_ID) = oRequest("EmployeeID").Item
				lErrorNumber = GetBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
			End If
		End If
		If lErrorNumber = 0 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
			Response.Write "<FORM NAME=""BanamexCensusFrm"" ID=""BanamexCensusFrm"" ACTION=""Catalogs.asp"" METHOD=""POST"" >"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BanamexCensus"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""ActionHdn"" VALUE=""2"" />"
				If Len(oRequest("AddNew").Item) = 0 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aBanamexCensusComponent(N_EMPLOYEE_ID) & """ />"
				End If
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""450px"">"
				If Len(oRequest("AddNew").Item) = 0 Then
					Response.Write "<TR>"
						Response.Write "<TD WIDTH=""3500""><FONT FACE=""Arial"" SIZE=""2"">Empleado:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_EMPLOYEE_ID) & "</FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empleado:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE="""" SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				End If
					If Len(oRequest("AddNew").Item) = 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Versión:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""u_version"" ID=""u_versionTxt"" VALUE=""" & aBanamexCensusComponent(N_U_VERSION) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Versión:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""u_version"" ID=""u_versionTxt"" VALUE="""" SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					End If
					If Len(oRequest("Modify").Item) > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Paterno:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Materno:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aBanamexCensusComponent(S_EMPLOYEE_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""rfc"" ID=""rfcTxt"" VALUE=""" & aBanamexCensusComponent(S_RFC) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""curp"" ID=""curpTxt"" VALUE=""" & aBanamexCensusComponent(S_CURP) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Núm. Seguro Social:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberTxt"" VALUE=""" & aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aBanamexCensusComponent(N_BIRTH_DATE), "BirthDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Lugar de nacimiento:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""BirthState"" ID=""BirthStateCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "States", "StateID", "StateCode, StateName", "", "StateName", aBanamexCensusComponent(S_STATE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Genero:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""GenderShortName"" ID=""GenderShortNameTxt"" VALUE=""" & aBanamexCensusComponent(S_GENDER_SHORT_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de ingreso:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aBanamexCensusComponent(N_JOIN_DATE), "JoinDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha cotización:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aBanamexCensusComponent(N_COT_DATE), "CotDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salario:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_SALARY) & "<FONT /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">FOVISSSTE:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_FOVY) & "<FONT /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""PeriodID"" ID=""PeriodIDTxt"" VALUE=""" & aBanamexCensusComponent(N_PERIOD_ID) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Status"" ID=""StatusTxt"" VALUE=""" & aBanamexCensusComponent(N_STATUS_ID) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">abre-cierra:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ChangeFlag"" ID=""ChangeFlagTxt"" VALUE=""" & aBanamexCensusComponent(N_CHANGE_FLAG) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado civil:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""MaritalStatus"" ID=""MaritalStatusTxt"" VALUE=""" & aBanamexCensusComponent(N_MARITAL_STATUS_ID) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Domicilio:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Address"" ID=""AddressTxt"" VALUE=""" & aBanamexCensusComponent(S_ADDRESS) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Colonia:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Colony"" ID=""ColonyTxt"" VALUE=""" & aBanamexCensusComponent(S_COLONY) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ciudad:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""City"" ID=""CityTxt"" VALUE=""" & aBanamexCensusComponent(S_CITY) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código postal:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ZipCode"" ID=""ZipCodeTxt"" VALUE=""" & aBanamexCensusComponent(N_ZIP_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""State"" ID=""StateCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "States", "StateID", "StateCode, StateName", "", "StateName", aBanamexCensusComponent(S_STATE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombramiento:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Nombram"" ID=""NombramTxt"" VALUE=""" & aBanamexCensusComponent(N_NOMBRAM) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Afore:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Afore"" ID=""AforeTxt"" VALUE=""" & aBanamexCensusComponent(N_AFORE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave ICEFA:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ICEFA"" ID=""ICEFATxt"" VALUE=""" & aBanamexCensusComponent(N_ICEFA) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número interno de control:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ICNumber"" ID=""ICNumberTxt"" VALUE=""" & aBanamexCensusComponent(N_IC_NUMBER) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Motivo de baja:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""mot_baja"" ID=""mot_bajaTxt"" VALUE=""" & aBanamexCensusComponent(S_MOT_BAJA) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salavio V:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Salary_v"" ID=""Salary_vTxt"" VALUE=""" & aBanamexCensusComponent(N_SALARY_V) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sueldo Integrado:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""FullPay"" ID=""FullPayTxt"" VALUE=""" & aBanamexCensusComponent(N_FULL_PAY) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días laborados:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""WorkingDays"" ID=""WorkingDaysTxt"" VALUE=""" & aBanamexCensusComponent(N_WORKING_DAYS) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de incapacidad:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""InabilityDays"" ID=""InabilityDaysTxt"" VALUE=""" & aBanamexCensusComponent(N_INABILITY_DAYS) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de ausencia:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""AbsenceDays"" ID=""AbsenceDaysTxt"" VALUE=""" & aBanamexCensusComponent(N_ABSENCE_DAYS) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Contribuciones del empleado:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeContributions"" ID=""EmployeeContributionsTxt"" VALUE=""" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto de las contribuciones:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeContributionsAmount"" ID=""EmployeeContributionsAmountTxt"" VALUE=""" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "StartDate", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha fin:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "EndDate", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
					ElseIf Len(oRequest("Delete").Item) > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD WIDTH=""50%""><FONT FACE=""Arial"" SIZE=""2"">Apellido Paterno:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Materno:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_EMPLOYEE_NAME) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_RFC) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_CURP) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Núm. Seguro Social:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_BIRTH_DATE) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Lugar de nacimiento:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_BIRTH_STATE_ID) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Genero:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_GENDER_SHORT_NAME) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de ingreso:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_JOIN_DATE) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha COT:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_COT_DATE) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salario:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_SALARY) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">FOVISSSTE:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_FOVY) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_PERIOD_ID) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_STATUS_ID) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">abre-cierra:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_CHANGE_FLAG) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado civil:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_MARITAL_STATUS_ID) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Domicilio:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_ADDRESS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Colonia:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_COLONY) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ciudad:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_CITY) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código postal:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_ZIP_NAME) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_STATE) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombramiento:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_NOMBRAM) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Afore:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_AFORE) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave ICEFA:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_ICEFA) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número interno de control:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_IC_NUMBER) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Motivo de baja:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(S_MOT_BAJA) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salavio V:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_SALARY_V) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sueldo Integrado:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_FULL_PAY) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días laborados:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_WORKING_DAYS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de incapacidad:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_INABILITY_DAYS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de ausencia:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_ABSENCE_DAYS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Contribuciones del empleado:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto de las contribuciones:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) & "</TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha fin:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) & "</TD>"
						Response.Write "</TR>"
					ElseIf Len(oRequest("AddNew").Item) > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Paterno:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""50"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Materno:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE="""" SIZE=""20"" MAXLENGTH=""50"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""rfc"" ID=""rfcTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""15"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""curp"" ID=""curpTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Núm. Seguro Social:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "BirthDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Lugar de nacimiento:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""BirthState"" ID=""BirthStateCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "States", "StateID", "StateCode, StateName", "", "StateName", aBanamexCensusComponent(S_STATE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Genero:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""GenderShortName"" ID=""GenderShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Genders", "GenderID", "GenderName", "GenderID>-1", "GenderID", aBanamexCensusComponent(S_GENDER_SHORT_NAME), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de ingreso:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(lCurrentDate, "JoinDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha cotización:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(lCurrentDate, "CotDate", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salario:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Salary"" ID=""SalaryTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">FOVISSSTE:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Fovi"" ID=""FoviTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""PeriodID"" ID=""PeriodIDTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado civil:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""MaritalStatus"" ID=""MaritalStatusCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "MaritalStatus", "MaritalStatusID", "MaritalStatusName", "", "MaritalStatusID", aBanamexCensusComponent(N_MARITAL_STATUS_ID), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Domicilio:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Address"" ID=""AddressTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Colonia:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Colony"" ID=""ColonyTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ciudad:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""City"" ID=""CityTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código postal:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ZipCode"" ID=""ZipCodeTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado:</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""State"" ID=""StateCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "States", "StateID", "StateCode, StateName", "", "StateName", aBanamexCensusComponent(S_STATE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombramiento:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Nombram"" ID=""NombramTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Afore:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Afore"" ID=""AforeTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave ICEFA:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ICEFA"" ID=""ICEFATxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número interno de control:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""ICNumber"" ID=""ICNumberTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Motivo de baja:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""mot_baja"" ID=""mot_bajaTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salavio V:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""Salary_v"" ID=""Salary_vTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sueldo Integrado:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""FullPay"" ID=""FullPayTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días laborados:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""WorkingDays"" ID=""WorkingDaysTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de incapacidad:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""InabilityDays"" ID=""InabilityDaysTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días de ausencia:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""AbsenceDays"" ID=""AbsenceDaysTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Contribuciones del empleado:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeContributions"" ID=""EmployeeContributionsTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto de las contribuciones:</FONT></TD><TD><INPUT TYPE=""TEXT"" NAME=""EmployeeContributionsAmount"" ID=""EmployeeContributionsAmountTxt"" VALUE="""" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "StartDate", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha fin:</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "EndDate", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
					End If
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR />"

				If aBanamexCensusComponent(N_ID_PROFILE) = -1 Then
					If aLoginComponent(N_USER_ID_FOR_CENSUS_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
				ElseIf Len(oRequest("Delete").Item) > 0 Then
					If aLoginComponent(N_USER_ID_FOR_CENSUS_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""RemoveCensus"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" />"
				Else
					If aLoginComponent(N_USER_ID_FOR_CENSUS_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ModifyCensus"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				End If
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Catalogs.asp?Action=ProfessionalRiskMatrix'"" />"
				Response.Write "<BR /><BR />"
				Call DisplayWarningDiv("RemoveCatalogWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
			Response.Write "</FORM>"
		End If
	End If
	DisplayBanamexCensusForm = lErrorNumber
	Err.Clear
End Function

Function DisplayBanamexCensusList(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the Banamex Census
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aBanamexCensusComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBanamexCensusList"
	Dim sNames
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim asCompanies
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sQuery
	Dim lErrorNumber
	Dim iStartPage
	Dim iRecordCounter

	sQuery = "Select CompanyID, CompanyName From Companies"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asCompanies = oRecordset.GetRows()
	iRecordCounter = 1
	lErrorNumber = GetBanamexCensusList(oRequest, oADODBConnection, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("&nbsp;,Empresa,Empleado,RFC,CURP,Núm. Seguro Social,Apellido paterno,Apellido materno,nombre,Centro de trabajo,Fecha de nacimiento,Estado,Genero,Fecha de contratación,Fecha cot,Sueldo,Fovissste,Periodo,Estatus,Abre-cierra,Estado civil,Domicilio,Colonia,Ciudad,Código postal,Estado,Nombramiento,Afore,Código ICEFA,Núm. Interno de control,Motivo baja,Salario V,Salario integrado,Días laborados,Días inhabiles,Días de ausencia,Contribución de empleado,Monto de la contribución,Fecha de inicio,Fecha fin,Usuario,Última actualización,Comentarios,Acciones", ",", -1, vbBinaryCompare)
			asCellWidths = Split("50,100,100,100,100,250,250,250,150,150,150,150,150,150,150,150,100,100,100,250,250,200,150,100,100,250,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,250", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,,,,,,CENTER,,CENTER,CENTER,CENTER,RIGHT,CENTER,CENTER,CENTER,CENTER,,,,,,,,CENTER,,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,,RIGHT,CENTER,CENTER,,CENTER,", ",", -1, vbBinaryCompare)
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
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(asCompanies(1,oRecordset.Fields("u_version").Value)))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & ""
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CT").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("BirthDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BirthState").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GenderShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("JoinDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("CotDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Salary").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Fovi").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Period").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Status").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ChangeFlag").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("MaritalStatusID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Address").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Colony").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("City").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZipZone").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("State").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nombram").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Afore").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ICEFA").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ICNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("mot_baja").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("Salary_v").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("FullPay").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("InabilityDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceDays").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeContributions").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(oRecordset.Fields("EmployeeContributionsAmount").Value, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("LastUpdateDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("EmployeeID").Value) <> 0 Then
							sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=BanamexCensus&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Modify=1"">"
							'sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=ProfessionalRisktable&BranchID=" & CStr(oRecordset.Fields("BranchID").Value) & "&CenterTypeID=" & CStr(oRecordset.Fields("CenterTypeID").Value) & "&PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & "&ServiceID=" & CStr(oRecordsetFields("ServiceID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_PERMISSIONS) = N_DELETE_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""Catalogs.asp?Action=BanamexCensus&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Delete=1"">"
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
	DisplayBanamexCensusList = lErrorNumber
	Err.Clear
End Function

Function DisplayDeletedHistoryList(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about previous deleted 
'		  employees from Banamex Census
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDeletedHistoryList"
	Dim sNames
	Dim oRecordset
	Dim asCompanies
	Dim asStates
	Dim asMaritalStatus
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sQuery
	Dim iRecordCounter
	
	sQuery = "Select CompanyID, CompanyName From Companies"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asCompanies = oRecordset.GetRows()
	sQuery = "Select StateID, StateName From States"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asStates = oRecordset.GetRows()
	sQuery = "Select MaritalStatusID, MaritalStatusName From MaritalStatus"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	asMaritalStatus = oRecordset.GetRows()
	iRecordCounter = 1
	sErrorDescription = "No se pudo obtener la información requerida"
	lErrorNumber = GetBanamexCensusList(oRequest, oADODBConnection, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se encontraron registros con los criterios indicados."
		lErrorNumber = -1
		If Not oRecordset.EOF Then
			lErrorNumber = 0
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""350"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Empresa,RFC,CURP,Núm. Seguro Social,Apellido paterno,Apellido materno,Nombre,CT,Fecha de Nacimiento,Estado de nacimiento,Género,Estado civil,Domicilio,Colonia,Ciudad,Código postal,Estado,Nombramiento,Núm. Empleado,ICEFA,Afore,Fecha de contratación,Fecha de cotización,Última actualización", ",", -1, vbBinaryCompare)
			asCellWidths = Split("150,100,100,100,150,150,150,100,150,150,100,100,150,150,150,100,150,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
			asCellAlignments = Split(" ,,,,,,,CENTER,,,CENTER,,,,,CENTER,,CENTER,,,,,,", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = ""
				sRowContents = CleanStringForHTML(CStr(asCompanies(1,oRecordset.Fields("u_version").Value)))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & ""
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CT").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("BirthDate").Value), -1, -1, -1)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(asStates(1,oRecordset.Fields("BirthState").Value)))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GenderShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(asMaritalStatus(1,oRecordset.Fields("MaritalStatusID").Value)))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Address").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Colony").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("City").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZipZone").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("State").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nombram").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ICEFA").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Afore").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("JoinDate").Value), -1, -1, -1)
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1)
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("LastUpdateDate").Value), -1, -1, -1)
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

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayDeletedHistoryList = lErrorNumber
	Err.Clear
End Function

Function CompareSarCensus(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To compare the payrolls agains the census and
'		  get new and deleted employees and information changes
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CompareSarCensus"
	Dim iIndex
	Dim lErrorNumber
	Dim oRecordset
	Dim sDate
	Dim sTColumns
	Dim asTColumns
	Dim sDocumentName
	Dim sFilePath
	Dim sFileName
	Dim sQuery
	Dim sRowContents

	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_Faltantes_SAR_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_Faltantes_SAR_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()

	sTColumns = "Empresa, Núm. Empleado, RFC, CURP, Núm. Seguro Social, Apellido Paterno, Apellido Materno, " & _
				"Centro de Trabajo, Fecha de nacimiento, Estado, Genero, Fecha de ingreso, Fecha de cotización, " & _
				"Salario, Fovi, Periodo, Estatus, Abre-Cierra, Estado civil, Dirección, Colonia, Ciudad, " & _
				"Código postal, Estado, Nombramiento, Afore, ICEFA, Control Interno, Motivo de Baja, Salario_V, " & _
				"Salario Integrado, Días Laborados, Incapacidades, Ausencias, Contribuciones del empleado, Monto, Fecha de inicio, Fecha fin"
	asTColumns = Split(sTColumns,",")
	
	sQuery = "Select * From Dm_Padron_Banamex Where (Status = 2) And (EmployeeID Not In (Select EmployeeID From Dm_Padron_banamex_Nuevo))"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_Altas_Faltantes_" & asTColumns(iIndex) & ".xls"
		sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
		For iIndex = 0 To UBound(asTColumns)
			sRowContents = sRowContents & "<TD>" & asTColumns(iIndex) & "</TD>"
		Next
		sRowContents = sRowContents & "</TR>"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		Do While Not oRecordset.EOF
			sRowContents = "<TR>"
			For iIndex = 0 To UBound(asTColumns)
				sRowContents = sRowContents & "<TD>" & oRecordset.Fields(iIndex) & "</TD>"
			Next
			sRowContents = sRowContents & "</ TR>"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			oRecordset.MoveNext
		Loop
		sRowContents = "</ TABLE>"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
	End If
	sQuery = "Select * From Dm_Padron_Banamex Where (Status = 3) And (EmployeeID In (Select EmployeeID From Dm_Padron_banamex_Nuevo))"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BanamexCensusComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		sDocumentName = sFilePath & "Rep_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & "_Bajas_Faltantes_" & asTColumns(iIndex) & ".xls"
		sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
		For iIndex = 0 To UBound(asTColumns)
			sRowContents = sRowContents & "<TD>" & asTColumns(iIndex) & "</TD>"
		Next
		sRowContents = sRowContents & "</TR>"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
		Do While Not oRecordset.EOF
			sRowContents = "<TR>"
			For iIndex = 0 To UBound(asTColumns)
				sRowContents = sRowContents & "<TD>" & oRecordset.Fields(iIndex) & "</TD>"
			Next
			sRowContents = sRowContents & "</ TR>"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			oRecordset.MoveNext
		Loop
		sRowContents = "</ TABLE>"
		lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
		sErrorDescription = "No se pudieron guardar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, " & sDate & ", '" & CATALOG_SEPARATOR & "', '', '', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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

	CompareSarCensus = lErrorNumber
	Err.Clear
End Function
%>