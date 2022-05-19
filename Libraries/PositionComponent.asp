<%
Const N_ID_POSITION = 1
Const S_SHORT_NAME_POSITION = 2
Const S_NAME_POSITION = 3
Const S_LONG_NAME_POSITION = 4
Const S_DESCRIPTION_POSITION = 5
Const N_START_DATE_POSITION = 6
Const N_END_DATE_POSITION = 7
Const N_EMPLOYEE_TYPE_ID_POSITION = 8
Const N_POSITION_TYPE_ID_POSITION = 9
Const N_COMPANY_ID_POSITION = 10
Const N_CLASSIFICATION_ID_POSITION = 11
Const N_GROUP_GRADE_LEVEL_ID_POSITION = 12
Const N_INTEGRATION_ID_POSITION = 13
Const N_LEVEL_ID_POSITION = 14
Const N_BRANCH_ID_POSITION = 15
Const N_SUB_BRANCH_ID_POSITION = 16
Const N_HIERARCHY_ID_POSITION = 17
Const N_GENERIC_POSITION_ID_POSITION = 18
Const D_WORKING_HOURS_POSITION = 19
Const N_STRATEGIC_POSITION = 20
Const N_NOMINATION_POSITION = 21
Const N_STATUS_ID_POSITION = 22
Const N_ACTIVE_POSITION = 23
Const N_DEPRECIATED_POSITION = 24
Const N_ECONOMICZONE = 25
Const S_COMMENTS = 26
Const D_AUTHORIZED_JOBS = 27
Const N_APPLIED_DATE_POSITION = 28

Const S_QUERY_CONDITION_POSITION = 29
Const B_CHECK_FOR_DUPLICATED_POSITION = 30
Const B_IS_DUPLICATED = 31
Const B_POSITION_COMPONENT_INITIALIZED = 32
Const S_FILTER_CONDITION_POSITION = 33

Const N_POSITION_COMPONENT_SIZE = 33
Dim aPositionComponent()
Redim aPositionComponent(N_POSITION_COMPONENT_SIZE)

Function InitializePositionComponent(oRequest, aPositionComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Position
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aPositionComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePositionComponent"
	Redim Preserve aPositionComponent(N_POSITION_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aPositionComponent(N_ID_POSITION)) Then
		If Len(oRequest("PositionID").Item) > 0 Then
			aPositionComponent(N_ID_POSITION) = CLng(oRequest("PositionID").Item)
		Else
			aPositionComponent(N_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(S_SHORT_NAME_POSITION)) Then
		If Len(oRequest("PositionShortName").Item) > 0 Then
			aPositionComponent(S_SHORT_NAME_POSITION) = CStr(oRequest("PositionShortName").Item)
		Else
			aPositionComponent(S_SHORT_NAME_POSITION) = ""
		End If
	End If
	aPositionComponent(S_SHORT_NAME_POSITION) = Left(aPositionComponent(S_SHORT_NAME_POSITION), 10)

	If IsEmpty(aPositionComponent(S_NAME_POSITION)) Then
		If Len(oRequest("PositionName").Item) > 0 Then
			aPositionComponent(S_NAME_POSITION) = CStr(oRequest("PositionName").Item)
		Else
			aPositionComponent(S_NAME_POSITION) = ""
		End If
	End If
	aPositionComponent(S_NAME_POSITION) = Left(aPositionComponent(S_NAME_POSITION), 255)

	If IsEmpty(aPositionComponent(S_LONG_NAME_POSITION)) Then
		If Len(oRequest("PositionLongName").Item) > 0 Then
			aPositionComponent(S_LONG_NAME_POSITION) = CStr(oRequest("PositionLongName").Item)
		Else
			aPositionComponent(S_LONG_NAME_POSITION) = aPositionComponent(S_NAME_POSITION)
		End If
	End If
	aPositionComponent(S_LONG_NAME_POSITION) = Left(aPositionComponent(S_LONG_NAME_POSITION), 2000)

	If IsEmpty(aPositionComponent(S_DESCRIPTION_POSITION)) Then
		If Len(oRequest("PositionDescription").Item) > 0 Then
			aPositionComponent(S_DESCRIPTION_POSITION) = CStr(oRequest("PositionDescription").Item)
		Else
			aPositionComponent(S_DESCRIPTION_POSITION) = " "
		End If
	End If
	aPositionComponent(S_DESCRIPTION_POSITION) = Left(aPositionComponent(S_DESCRIPTION_POSITION), 2000)

	If IsEmpty(aPositionComponent(N_START_DATE_POSITION)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aPositionComponent(N_START_DATE_POSITION) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDateYear").Item) > 0 Then
			aPositionComponent(N_START_DATE_POSITION) = CLng(oRequest("StartDateYear").Item & Right(("0" & oRequest("StartDateMonth").Item), Len("00")) & Right(("0" & oRequest("StartDateDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aPositionComponent(N_START_DATE_POSITION) = CLng(oRequest("StartDate").Item)
		Else
			aPositionComponent(N_START_DATE_POSITION) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aPositionComponent(N_END_DATE_POSITION)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aPositionComponent(N_END_DATE_POSITION) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDateYear").Item) > 0 Then
			aPositionComponent(N_END_DATE_POSITION) = CLng(oRequest("EndDateYear").Item & Right(("0" & oRequest("EndDateMonth").Item), Len("00")) & Right(("0" & oRequest("EndDateDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aPositionComponent(N_END_DATE_POSITION) = CLng(oRequest("EndDate").Item)
		Else
			aPositionComponent(N_END_DATE_POSITION) = 30000000
		End If
	End If
	If CLng(aPositionComponent(N_END_DATE_POSITION)) = 0 Then aPositionComponent(N_END_DATE_POSITION) = 30000000

	If IsEmpty(aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION)) Then
		If Len(oRequest("EmployeeTypeID").Item) > 0 Then
			aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) = CInt(oRequest("EmployeeTypeID").Item)
		Else
			aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_POSITION_TYPE_ID_POSITION)) Then
		If Len(oRequest("PositionTypeID").Item) > 0 Then
			aPositionComponent(N_POSITION_TYPE_ID_POSITION) = CInt(oRequest("PositionTypeID").Item)
		Else
			aPositionComponent(N_POSITION_TYPE_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_COMPANY_ID_POSITION)) Then
		If Len(oRequest("CompanyID").Item) > 0 Then
			aPositionComponent(N_COMPANY_ID_POSITION) = CInt(oRequest("CompanyID").Item)
		Else
			aPositionComponent(N_COMPANY_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_CLASSIFICATION_ID_POSITION)) Then
		If Len(oRequest("ClassificationID").Item) > 0 Then
			aPositionComponent(N_CLASSIFICATION_ID_POSITION) = CInt(oRequest("ClassificationID").Item)
		Else
			aPositionComponent(N_CLASSIFICATION_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION)) Then
		If Len(oRequest("GroupGradeLevelID").Item) > 0 Then
			aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) = CInt(oRequest("GroupGradeLevelID").Item)
		Else
			aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_INTEGRATION_ID_POSITION)) Then
		If Len(oRequest("IntegrationID").Item) > 0 Then
			aPositionComponent(N_INTEGRATION_ID_POSITION) = CInt(oRequest("IntegrationID").Item)
		Else
			aPositionComponent(N_INTEGRATION_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_LEVEL_ID_POSITION)) Then
		If Len(oRequest("LevelID").Item) > 0 Then
			aPositionComponent(N_LEVEL_ID_POSITION) = CInt(oRequest("LevelID").Item)
		Else
			aPositionComponent(N_LEVEL_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_BRANCH_ID_POSITION)) Then
		If Len(oRequest("BranchID").Item) > 0 Then
			aPositionComponent(N_BRANCH_ID_POSITION) = CInt(oRequest("BranchID").Item)
		Else
			aPositionComponent(N_BRANCH_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_SUB_BRANCH_ID_POSITION)) Then
		If Len(oRequest("SubBranchID").Item) > 0 Then
			aPositionComponent(N_SUB_BRANCH_ID_POSITION) = CInt(oRequest("SubBranchID").Item)
		Else
			aPositionComponent(N_SUB_BRANCH_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_HIERARCHY_ID_POSITION)) Then
		If Len(oRequest("HierarchyID").Item) > 0 Then
			aPositionComponent(N_HIERARCHY_ID_POSITION) = CInt(oRequest("HierarchyID").Item)
		Else
			aPositionComponent(N_HIERARCHY_ID_POSITION) = 0
		End If
	End If

	If IsEmpty(aPositionComponent(N_GENERIC_POSITION_ID_POSITION)) Then
		If Len(oRequest("GenericPositionID").Item) > 0 Then
			aPositionComponent(N_GENERIC_POSITION_ID_POSITION) = CInt(oRequest("GenericPositionID").Item)
		Else
			aPositionComponent(N_GENERIC_POSITION_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(D_WORKING_HOURS_POSITION)) Then
		If Len(oRequest("WorkingHours").Item) > 0 Then
			aPositionComponent(D_WORKING_HOURS_POSITION) = CDbl(oRequest("WorkingHours").Item)
		Else
			aPositionComponent(D_WORKING_HOURS_POSITION) = 0
		End If
	End If

	If IsEmpty(aPositionComponent(N_STRATEGIC_POSITION)) Then
		If Len(oRequest("Strategic").Item) > 0 Then
			aPositionComponent(N_STRATEGIC_POSITION) = CInt(oRequest("Strategic").Item)
		Else
			aPositionComponent(N_STRATEGIC_POSITION) = 1
		End If
	End If

	If IsEmpty(aPositionComponent(N_NOMINATION_POSITION)) Then
		If Len(oRequest("Nomination").Item) > 0 Then
			aPositionComponent(N_NOMINATION_POSITION) = CInt(oRequest("Nomination").Item)
		Else
			aPositionComponent(N_NOMINATION_POSITION) = 1
		End If
	End If

	If IsEmpty(aPositionComponent(N_STATUS_ID_POSITION)) Then
		If Len(oRequest("StatusID").Item) > 0 Then
			aPositionComponent(N_STATUS_ID_POSITION) = CInt(oRequest("StatusID").Item)
		Else
			aPositionComponent(N_STATUS_ID_POSITION) = -1
		End If
	End If

	If IsEmpty(aPositionComponent(N_ACTIVE_POSITION)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aPositionComponent(N_ACTIVE_POSITION) = CInt(oRequest("Active").Item)
		Else
			aPositionComponent(N_ACTIVE_POSITION) = 1
		End If
	End If

	If IsEmpty(aPositionComponent(N_DEPRECIATED_POSITION)) Then
		If Len(oRequest("Depreciated").Item) > 0 Then
			aPositionComponent(N_DEPRECIATED_POSITION) = CInt(oRequest("Depreciated").Item)
		Else
			aPositionComponent(N_DEPRECIATED_POSITION) = 0
		End If
	End If

	If IsEmpty(aPositionComponent(N_ECONOMICZONE)) Then
		If Len(oRequest("EconomicZoneID").Item) > 0 Then
			aPositionComponent(N_ECONOMICZONE) = CInt(oRequest("EconomicZoneID").Item)
		Else
			aPositionComponent(N_ECONOMICZONE) = 0
		End If
	End If

	If IsEmpty(aPositionComponent(S_COMMENTS)) Then
		If Len(oRequest("Comments").Item) > 0 Then
			aPositionComponent(S_COMMENTS) = CStr(oRequest("Comments").Item)
		Else
			aPositionComponent(S_COMMENTS) = " "
		End If
	End If
    If IsEmpty(aPositionComponent(D_AUTHORIZED_JOBS)) Then
		If Len(oRequest("AuthorizedJobs").Item) > 0 Then
			aPositionComponent(D_AUTHORIZED_JOBS) = CInt(oRequest("AuthorizedJobs").Item)
		Else
			aPositionComponent(D_AUTHORIZED_JOBS) = 0
		End If
	End If
	If IsEmpty(aPositionComponent(N_APPLIED_DATE_POSITION)) Then
		If Len(oRequest("AppliedDate").Item) > 0 Then
			aPositionComponent(N_APPLIED_DATE_POSITION) = CLng(oRequest("AppliedDate").Item)
		Else
			aPositionComponent(N_APPLIED_DATE_POSITION) = -1
		End If
	End If

	aPositionComponent(S_QUERY_CONDITION_POSITION) = ""
	aPositionComponent(B_CHECK_FOR_DUPLICATED_POSITION) = True
	aPositionComponent(B_IS_DUPLICATED_POSITION) = False

	aPositionComponent(B_POSITION_COMPONENT_INITIALIZED) = True
	InitializePositionComponent = Err.number
	Err.Clear
End Function

Function AddPosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPosition"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
    Dim lPositionID
    Dim lPositionStartDate

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePositionComponent(oRequest, aPositionComponent)
	End If

	If Not CheckExistencyOfPosition(aPositionComponent, lPositionID, lPositionStartDate, sErrorDescription) Then
		lErrorNumber = L_ERR_DUPLICATED_RECORD
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aPositionComponent(N_ID_POSITION) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Positions", "PositionID", "", 1, aPositionComponent(N_ID_POSITION), sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la información del nuevo registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Positions (PositionID, PositionShortName, PositionName, PositionLongName, PositionDescription, StartDate, EndDate, EmployeeTypeID, PositionTypeID, CompanyID, ClassificationID, GroupGradeLevelID, IntegrationID, LevelID, BranchID, SubBranchID, HierarchyID, GenericPositionID, WorkingHours, Strategic, Nomination, StatusID, Active, Depreciated,EconomicZoneID, Comments, AuthorizedJobs) Values (" & aPositionComponent(N_ID_POSITION) & ", '" & Replace(aPositionComponent(S_SHORT_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_LONG_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_DESCRIPTION_POSITION), "'", "´") & "', " & aPositionComponent(N_START_DATE_POSITION) & ", " & aPositionComponent(N_END_DATE_POSITION) & ", " & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ", " & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ", " & aPositionComponent(N_COMPANY_ID_POSITION) & ", " & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ", " & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ", " & aPositionComponent(N_INTEGRATION_ID_POSITION) & ", " & aPositionComponent(N_LEVEL_ID_POSITION) & ", " & aPositionComponent(N_BRANCH_ID_POSITION) & ", " & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ", " & aPositionComponent(N_HIERARCHY_ID_POSITION) & ", " & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ", " & aPositionComponent(D_WORKING_HOURS_POSITION) & ", " & aPositionComponent(N_STRATEGIC_POSITION) & ", " & aPositionComponent(N_NOMINATION_POSITION) & ", " & aPositionComponent(N_STATUS_ID_POSITION) & ", " & aPositionComponent(N_ACTIVE_POSITION) & ", " & aPositionComponent(N_DEPRECIATED_POSITION) & "," & aPositionComponent(N_ECONOMICZONE) & ", '" & Replace(aPositionComponent(S_COMMENTS), "'", "") & "', " & aPositionComponent(D_AUTHORIZED_JOBS) &")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If
	AddPosition = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPosition(aPositionComponent, lPositionID, lPositionStartDate, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aPositionComponent
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPosition"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sPositionCrossType

	sErrorDescription = "No se pudo verificar la existencia de puestos en la base de datos."
	sQuery = "Select * From Positions Where (PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And (StartDate<>" & aPositionComponent(N_START_DATE_POSITION) & ")" & _
            " And (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "')" & _
			" And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ")" & _
			" And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ")" & _
			" And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ")" & _
			" And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ")" & _
			" And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ")" & _
			" And (((StartDate >= " &  aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate <= " &  aPositionComponent(N_END_DATE_POSITION) & "))" & _
			" Or ((EndDate >= " &  aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate <= " &  aPositionComponent(N_END_DATE_POSITION) & "))" & _
			" Or ((EndDate >= " &  aPositionComponent(N_START_DATE_POSITION) & ") And (StartDate <= " &  aPositionComponent(N_END_DATE_POSITION) & "))) Order By StartDate Desc"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If aPositionComponent(N_START_DATE_POSITION) <> CLng(oRecordset.Fields("StartDate").Value) Then
				lErrorNumber = GetPositionCrossType(oADODBConnection, aPositionComponent, sPositionCrossType, lPositionID, lPositionStartDate, sErrorDescription)
				If lErrorNumber = 0 Then
					Select Case sPositionCrossType
						Case "Left"
							aPositionComponent(N_STATUS_ID_POSITION) = 0
					        CheckExistencyOfPosition = False
                            sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
						Case "Right"
							aPositionComponent(N_STATUS_ID_POSITION) = -1
                            CheckExistencyOfPosition = False
                            sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
						Case "Inner"
							aPositionComponent(N_STATUS_ID_POSITION) = -2
                            CheckExistencyOfPosition = False
                            sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
						Case "Cross"
							aPositionComponent(N_STATUS_ID_POSITION) = -3
                            CheckExistencyOfPosition = False
                            sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
					End Select
				Else
					sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
					CheckExistencyOfPosition = False
				End If
			Else
				sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que existe uno registrado con la misma fecha de inicio"
				CheckExistencyOfPosition = False
			End If
		Else
			aPositionComponent(N_STATUS_ID_POSITION) = 0
			CheckExistencyOfPosition = True
		End If
	Else
		sErrorDescription = "No se puede agregar el puesto " & aPositionComponent(S_SHORT_NAME_POSITION) & " con fecha de inicio " & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & " debido a que hubo error al verificar si existe uno registrado en el periodo indicado"
		CheckExistencyOfPosition = False
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function DisplayPositionForm(oRequest, oADODBConnection, sAction, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aPositionComponent
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionForm"
	Dim lErrorNumber
	Dim sFilter

	sFilter = ""

    sFilter = "ApplyFilter=1&PositionShortNameFilter=" & CStr(oRequest("PositionShortNameFilter").Item) & "&GroupGradeLevelIDFilter=" & CStr(oRequest("GroupGradeLevelIDFilter").Item) & "&EmployeeTypeIDFilter=" & CStr(oRequest("EmployeeTypeIDFilter").Item)
    If aPositionComponent(N_ID_POSITION) <> -1 Then
		lErrorNumber = GetPositionn(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine


			Response.Write "function TotalDays(iDay, iMonth, iYear){" & vbNewLine
				Response.Write "iMonth = (iMonth + 9) % 12;" & vbNewLine
				Response.Write "iYear = iYear - Math.floor(iMonth/10);" & vbNewLine
				Response.Write "return (365 * iYear + Math.floor(iYear/4) - Math.floor(iYear/100) + Math.floor(iYear/400) + Math.floor((iMonth * 306 + 5)/10) + iDay - 1)" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "function GetDifferenceBetweenDates(iDay1, iMonth1, iYear1, iDay2, iMonth2, iYear2){" & vbNewLine
				Response.Write "return TotalDays(iDay2, iMonth2, iYear2) - TotalDays(iDay1, iMonth1, iYear1)" & vbNewLine
			Response.Write "}" & vbNewLine


			Response.Write "function CheckPositionFields(oForm) {" & vbNewLine
                Response.Write "alert('CheckPositionFields...');" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.PositionShortName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir la clave del puesto.');" & vbNewLine
						Response.Write "oForm.PositionShortName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.PositionName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del puesto.');" & vbNewLine
						Response.Write "oForm.PositionName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.PositionLongName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre detallado del puesto.');" & vbNewLine
						Response.Write "oForm.PositionLongName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "oForm.ClassificationID.value = oForm.ClassificationID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.ClassificationID, 'el campo \'Clasificación\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, -1, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.IntegrationID.value = oForm.IntegrationID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.IntegrationID, 'el campo \'Integración\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, -1, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.HierarchyID.value = oForm.HierarchyID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.TaxMax, 'el monto máximo gravable del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.WorkingHours.value = oForm.WorkingHours.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.WorkingHours, 'el campo \'Horas laboradas\'', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 24))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((parseInt('1' + oForm.EndDay.value) - 100) + ('1' + parseInt(oForm.EndMonth.value) - 100) + parseInt(oForm.EndYear.value)) > 0 ) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.EndDay.value) - 100) * (parseInt('1' + oForm.EndMonth.value) - 100) * parseInt(oForm.EndYear.value) > 0 ) {" & vbNewLine
						Response.Write "if (((parseInt('1' + oForm.StartDay.value) - 100) + ((parseInt('1' + oForm.StartMonth.value) -100) * 100) + parseInt(oForm.StartYear.value) * 10000) > ((parseInt('1' + oForm.EndDay.value) - 100) + ((parseInt('1' + oForm.EndMonth.value) - 100) * 100) + parseInt(oForm.EndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del registro de crédito del empleado.');" & vbNewLine
							Response.Write "oForm.EndDay.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else" & vbNewLine
					Response.Write "{" & vbNewLine
						Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckPositionFields" & vbNewLine
			Response.Write "function ShowAmountFields(sValue, sFieldsName) {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (sValue == '1')" & vbNewLine
						Response.Write "ShowDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
					Response.Write "else" & vbNewLine
						Response.Write "HideDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine
			Response.Write "function SetTaxValue() {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine
				Response.Write "var i;" & vbNewLine
				Response.Write "var j;" & vbNewLine
					Response.Write "for (i=0;i<oForm.ForTax.length;i++){" & vbNewLine
						Response.Write "if (oForm.ForTax[i].checked)" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "for (j=0;j<oForm.IsDeduction.length;j++){" & vbNewLine
						Response.Write "if (oForm.IsDeduction[j].checked)" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (oForm.IsDeduction[j].value == 0) {" & vbNewLine
						Response.Write "if (oForm.ForTax[i].value == 1) {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=100;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=0;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else {" & vbNewLine
						Response.Write "if (oForm.ForTax[i].value == 1) {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=0;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=100;" & vbNewLine

						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine

			Response.Write "function CheckCompletedFields(oForm){" &vbNewLine
				Response.Write "if(oForm){" &vbNewLine
					Response.Write "if(oForm.CompanyID.value<0){" &vbNewLine
						Response.Write "alert('Favor de seleccionar la empresa del puesto');" &vbNewLine
						Response.Write "oForm.CompanyID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.EmployeeTypeID.value<0){" &vbNewLine
						Response.Write "alert('Favor de seleccionar el tipo de tabulador del puesto');" &vbNewLine
						Response.Write "oForm.EmployeeTypeID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.PositionTypeID.value<0){" &vbNewLine
						Response.Write "alert('Favor de seleccionar el tipo de puesto del puesto');" &vbNewLine
						Response.Write "oForm.PositionTypeID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.PositionShortName.value.length==0){" &vbNewLine
						Response.Write "alert('Ingrese el código del puesto');" &vbNewLine
						Response.Write "oForm.PositionShortName.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.PositionName.value.length==0){" &vbNewLine
						Response.Write "alert('Favor de introducir el nombre del puesto');" &vbNewLine
						Response.Write "oForm.PositionName.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine

					Response.Write "if (((parseInt('1' + oForm.EndDay.value) - 100) + ('1' + parseInt(oForm.EndMonth.value) - 100) + parseInt(oForm.EndYear.value)) > 0 ) {" &vbNewLine
                        Response.Write "if ((parseInt('1' + oForm.EndDay.value) - 100) * (parseInt('1' + oForm.EndMonth.value) - 100) * parseInt(oForm.EndYear.value) > 0 ) {" &vbNewLine
                            If Len(oRequest("Change").Item) > 0 Then
                                Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EndDay.value), parseInt(oForm.EndMonth.value), parseInt(oForm.EndYear.value), parseInt(oForm.AppliedDate.value.substr(6,2)), parseInt(oForm.AppliedDate.value.substr(4,2)), parseInt(oForm.AppliedDate.value.substr(0,4))) > 0) {" & vbNewLine
				                    Response.Write "alert('La fecha de fin del registro del puesto no puede ser menor que la fecha de la aplicación.');" & vbNewLine
				                    Response.Write "oForm.EndDay.focus();" & vbNewLine
				                    Response.Write "return false;" & vbNewLine
				                Response.Write "}" & vbNewLine
                            Else
                                Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EndDay.value), parseInt(oForm.EndMonth.value), parseInt(oForm.EndYear.value), parseInt(oForm.StartDay.value), parseInt(oForm.StartMonth.value), parseInt(oForm.StartYear.value)) > 0) {" & vbNewLine
								    Response.Write "alert('La fecha de fin del registro del puesto no puede ser menor que la fecha de inicio.');" & vbNewLine
								    Response.Write "oForm.EndDay.focus();" &vbNewLine
								    Response.Write "return false;" &vbNewLine
							    Response.Write "}" &vbNewLine
                            End If
                        Response.Write "}" &vbNewLine
						Response.Write "else{" &vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del movimiento');" &vbNewLine
							Response.Write "return false;" &vbNewLine
						Response.Write "}" &vbNewLine
					Response.Write "}" &vbNewLine

					Response.Write "if(oForm.GroupGradeLevelID.value<0 && document.getElementById('GroupGradeLevelCmbID').disabled==false){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente el Grupo, grado, nivel del puesto');" &vbNewLine
						Response.Write "oForm.GroupGradeLevelID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.WorkingHours.value<=0){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente la jornada laboral del puesto');" &vbNewLine
						Response.Write "oForm.WorkingHours.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.ClassificationID.value <=0 && document.getElementById('Classificationtxt').disabled==false ){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente la clasificación del puesto');" &vbNewLine
						Response.Write "oForm.ClassificationID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.IntegrationID.value<0 && document.getElementById('Integrationtxt').disabled==false){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente la integración del puesto');" &vbNewLine
						Response.Write "oForm.IntegrationID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "if(oForm.GenericPositionID.value<0 && document.getElementById('GenericPositionIDCmb').disabled==false){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente el puesto genérico del puesto');" &vbNewLine
						Response.Write "oForm.GenericPositionID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
                    Response.Write "if(oForm.EconomicZoneID.value<=0 && document.getElementById('EconomicZoneIDCmb').disabled==false){" &vbNewLine
						Response.Write"alert('Favor de introducir correctamente la zona económica del puesto');" &vbNewLine
						Response.Write "oForm.EconomicZoneID.focus();" &vbNewLine
						Response.Write "return false;" &vbNewLine
					Response.Write "}" &vbNewLine
					Response.Write "oForm.ClassificationID.value = oForm.ClassificationID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.ClassificationID, 'el campo \'Clasificación\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, -1, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.IntegrationID.value = oForm.IntegrationID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.IntegrationID, 'el campo \'Integración\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, -1, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.HierarchyID.value = oForm.HierarchyID.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.TaxMax, 'el monto máximo gravable del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "oForm.WorkingHours.value = oForm.WorkingHours.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.WorkingHours, 'el campo \'Horas laboradas\'', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 24))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" &vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "}  // End of CheckCompletedFields" & vbNewLine
			Response.Write "function enableFields(){" &vbNewLine
				Response.Write " var text = document.getElementById('EmployeeTypeIDCmb').options[document.getElementById('EmployeeTypeIDCmb').selectedIndex].value;" &vbNewLine
				Response.Write "if(text ==1){" & vbNewLine
                    Response.Write "document.getElementById('EconomicZoneIDCmb').selectedIndex = 0" &vbNewLine
					Response.Write "document.getElementById('Classificationtxt').disabled=false;" &vbNewLine
					Response.Write "document.getElementById('Integrationtxt').disabled=false;" &vbNewLine
					Response.Write "document.getElementById('GroupGradeLevelCmbID').disabled=false;" &vbNewLine
                    Response.Write "document.getElementById('GroupGradeLevelCmbID').disabled=false;" &vbNewLine
                    Response.Write "document.getElementById('EconomicZoneIDCmb').disabled=true;" &vbNewLine
                    Response.Write "document.getElementById('LevelIDCmb').disabled=true;" &vbNewLine
				Response.Write "}" &vbNewLine
                Response.Write "else if(text ==3){" & vbNewLine
                    Response.Write "document.getElementById('EconomicZoneIDCmb').selectedIndex = 0" &vbNewLine
                    Response.Write "document.getElementById('EconomicZoneIDCmb').disabled=true;" &vbNewLine
				Response.Write "}" &vbNewLine
				Response.Write "else{" & vbNewLine
                    Response.Write "document.getElementById('Classificationtxt').value = -1" &vbNewLine
                    Response.Write "document.getElementById('GroupGradeLevelCmbID').selectedIndex = 0" &vbNewLine
                    Response.Write "document.getElementById('Integrationtxt').value = -1" &vbNewLine
					Response.Write "document.getElementById('Classificationtxt').disabled=true;" &vbNewLine
					Response.Write "document.getElementById('Integrationtxt').disabled=true;" &vbNewLine
					Response.Write "document.getElementById('GroupGradeLevelCmbID').disabled=true;" &vbNewLine
                    Response.Write "document.getElementById('EconomicZoneIDCmb').disabled=false;" &vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" &vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
        Response.Write "<FORM NAME=""ConceptFrm"" ID=""ConceptFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckCompletedFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Positions"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aPositionComponent(N_ID_POSITION) & """ />"
			'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aPositionComponent(N_START_DATE_POSITION) & """ />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"

                Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"					
							Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1""> </OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(CompanyID>0) And (Active=1)", "CompanyShortName", aPositionComponent(N_COMPANY_ID_POSITION), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de tabulador:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onchange=""enableFields()"">"
								Response.Write "<OPTION VALUE=""-1""> </OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "(EmployeeTypeID>=0) And (EmployeeTypeID<7) And (Active=1)", "EmployeeTypeName", aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de puesto:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><SELECT NAME=""PositionTypeID"" ID=""PositionTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" >"
								Response.Write "<OPTION VALUE=""-1""> </OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionTypes", "PositionTypeID", "PositionTypeName", "(PositionTypeID>=0) And (PositionTypeID<7) And (Active=1)", "PositionTypeName", aPositionComponent(N_POSITION_TYPE_ID_POSITION), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PositionShortName"" ID=""PositionShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aPositionComponent(S_SHORT_NAME_POSITION)) & """ CLASS=""TextFields"""
					If Len(oRequest("Change").Item) > 0 Then
						Response.Write  " READONLY=""READONLY"" /></TD>"
					Else
						Response.Write " /></TD>"
					End If
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PositionName"" ID=""PositionNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aPositionComponent(S_NAME_POSITION)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"

                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						'Response.Write DisplayDateCombosUsingSerial(aPositionComponent(N_START_DATE_POSITION), "Start", N_START_YEAR, Year(Date) + 1, True, False)
                        If Len(oRequest("Change").Item) > 0 Then
				            Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aPositionComponent(N_START_DATE_POSITION), -1, -1, -1) & "</FONT>"
				            Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aPositionComponent(N_START_DATE_POSITION) & """ />"
                        Else
						    Response.Write DisplayDateCombosUsingSerial(aPositionComponent(N_START_DATE_POSITION), "Start", N_START_YEAR, Year(Date) + 1, True, False)
                        End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aPositionComponent(N_END_DATE_POSITION), "End", N_START_YEAR, Year(Date())+2, True, True)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"

				If Len(oRequest("Change").Item) > 0 Then
                    Response.Write "<TR>"
					    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación:&nbsp;</NOBR></FONT></TD>"
					    Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
							    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (PayrollTypeID=1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
					    Response.Write "</SELECT>&nbsp;"
					    Response.Write "</TD>"
				    Response.Write "</TR>"
                End If

                Response.Write "<TR>"
				    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						    Response.Write "<OPTION VALUE=""-1""> </OPTION>"
						    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelName", "(LevelID>-1) And (Active=1)", "LevelID", aPositionComponent(N_LEVEL_ID_POSITION), "", sErrorDescription)
					    Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelCmbID"" DISABLED=""TRUE"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-1""> </OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelShortName", "(GroupGradeLevelID>-1) And (Active=1)", "GroupGradeLevelName", aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
			    Response.Write "</TR>"

                Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Integración:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""Integrationtxt"" SIZE=""5"" DISABLED=""TRUE"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aPositionComponent(N_INTEGRATION_ID_POSITION)) & """ CLASS=""TextFields"" /></TD>"
						'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aPositionComponent(N_INTEGRATION_ID_POSITION)) & "</FONT></TD>"
				Response.Write "</TR>"

                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clasificación:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""Classificationtxt"" DISABLED=""TRUE""  SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & """ CLASS=""TextFields"" /></TD>"
					'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & "</FONT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Zona Económica:&nbsp;</B></FONT></TD>"
				    Response.Write "<TD><SELECT NAME=""EconomicZoneID"" ID=""EconomicZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""0""> </OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EconomicZones", "EconomicZoneID", "EconomicZoneName", "(EconomicZoneID>=0) And (Active=1)", "EconomicZoneName", aPositionComponent(N_ECONOMICZONE), "", sErrorDescription)
						Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Jornada:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHours"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aPositionComponent(D_WORKING_HOURS_POSITION)) & """ CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">Hrs.</FONT></TD>"
					'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aPositionComponent(D_WORKING_HOURS_POSITION)) & "</FONT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto Generico:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""GenericPositionID"" ID=""GenericPositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-1""> </OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GenericPositions", "GenericPositionID", "GenericPositionName", "(GenericPositionID>-1) And (Active=1)", "GenericPositionID", aPositionComponent(N_GENERIC_POSITION_ID_POSITION), "", sErrorDescription)
						Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Rama:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""BranchID"" ID=""BranchIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-1""> </OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Branches", "BranchID", "BranchName", "(BranchID>-1) And (Active=1)", "BranchID", aPositionComponent(N_BRANCH_ID_POSITION), "", sErrorDescription)
						Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sub-Rama:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""SubBranchID"" ID=""SubBranchIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-1""> </OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SubBranches", "SubBranchID", "SubBranchName", "(SubBranchID>-1) And (Active=1)", "SubBranchID", aPositionComponent(N_SUB_BRANCH_ID_POSITION), "", sErrorDescription)
		    			Response.Write "</SELECT></TD>"
			    Response.Write "</TR>"
                Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>No. de plazas authorizadas:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AuthorizedJobs"" ID=""AuthorizedJobs"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aPositionComponent(D_AUTHORIZED_JOBS)) & """ CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2""></FONT></TD>"
					'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aPositionComponent(D_WORKING_HOURS_POSITION)) & "</FONT></TD>"
				Response.Write "</TR>"
                Response.Write"<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Comentarios:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""60"" CLASS=""TextFields"">" & aPositionComponent(S_COMMENTS) &"</TEXTAREA></TD>"
				Response.Write"</TR>"

			Response.Write "</TABLE>"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "if (document.ConceptFrm.IsCredit.checked) {ShowDisplay(document.all['ReasonsDiv']) } else {HideDisplay(document.all['ReasonsDiv'])};" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
			'Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			'	Response.Write "ShowAmountFields(document.ConceptFrm.TaxQttyID.value, 'Tax');" & vbNewLine
			'	Response.Write "ShowAmountFields(document.ConceptFrm.ExemptQttyID.value, 'Exempt');" & vbNewLine
			'Response.Write "//--></SCRIPT>" & vbNewLine
			If aPositionComponent(N_ID_POSITION) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
            If Len(oRequest("ApplyFilter").Item) = 0 Then
                Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
            Else
                If aPositionComponent(N_ID_POSITION) <> -1 Then
			        Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("")  & "?" & sFilter & "'"" />"
                Else
                    Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
                End If
            End If
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayPositionForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPositionsTable(oRequest, oADODBConnection, bForExport, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To display the ConceptValues for Concepts
'Inputs:  oRequest, oADODBConnection, iSelectedTab, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionsTable"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sDate
	Dim sStartDate
	Dim sEndDate
	Dim sFilePath
	Dim lReportID
	Dim sTemp
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asColumnsTitles
	Dim asCellWidths
	Dim asCellAlignments
	Dim sColumnsTitles
	Dim sCellWidths
	Dim sCellAlignments
	Dim lErrorNumber
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim iRecordCounter
	Dim sLevelShortName
	Dim sFilter
	Dim sPage

	sFilter = ""
	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))

	sErrorDescription = "No se pudieron obtener los puestos registrados."
	lErrorNumber = GetPositions(oRequest, oADODBConnection, aPositionComponent, oRecordset, sErrorDescription)

	If Len(oRequest("ApplyFilter").Item) > 0 Then
		sFilter = "&ApplyFilter=1&StartForValueDay="& oRequest("StartForValueDay") &"&StartForValueMonth="& oRequest("StartForValueMonth") &"&StartForValueYear="& oRequest("StartForValueYear") &"&EndForValueDay="& oRequest("EndForValueDay") &"&EndForValueMonth="& oRequest("EndForValueMonth") &"&EndForValueYear="& oRequest("EndForValueYear") &"&PositionShortNameFilter=" & oRequest("PositionShortNameFilter") & "&GroupGradeLevelIDFilter=" & oRequest("GroupGradeLevelIDFilter") & "&EmployeeTypeIDFilter=" & oRequest("EmployeeTypeIDFilter")
	End If

	If Len(oRequest("StartPage"))>0 Then
		sPage = "&StartPage="& oRequest("StartPage") &""
	End If
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_CATALOG, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine

			If bForExport Then
                sColumnsTitles = "Empresa,Tipo de tabulador,Tipo de puesto,Código,Denominación,Nivel,Grupo-grado-nivel,Integración,Clasificación,Zona Económica,Jornada,Rama,Subrama,Jerarquía,Puesto Generico,No. Plazas Autorizadas, Fecha de inicio,Fecha de término,Comentarios"
            Else
                sColumnsTitles = "Empresa,Tipo de tabulador,Tipo de puesto,Código,Denominación,Nivel,Grupo-grado-nivel,Integración,Clasificación,Zona Económica,Jornada,Rama,Subrama,Jerarquía,Puesto Generico,No. Plazas Autorizadas, Fecha de inicio,Fecha de término"
            End If
			sCellWidths = "200,100,30,100,800,50,100,100,,,,,,"
			sCellAlignments = "CENTER,LEFT,CENTER,CENTER,LEFT,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,CENTER,LEFT"
			If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
				sColumnsTitles = "Acciones," & sColumnsTitles
				sCellWidths = "90," & sCellWidths
				sCellAlignments = "CENTER," & sCellAlignments
			End If
			asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
			asCellWidths = Split(sCellWidths, ",", -1, vbBinaryCompare)
			asCellAlignments = Split(sCellAlignments, ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			lCurrentPositionID = -2
			dTotal = 0
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
            sTextBegin =""
            sTextEnd=""
			iRecordCounter = 0
			Do While Not oRecordset.EOF
                sRowContents = ""    
                bContinue = False
				sBoldBegin = ""
				sBoldEnd = ""
				If StrComp(CStr(oRecordset.Fields("PositionID").Value), oRequest("PositionID").Item, vbBinaryCompare) = 0  And StrComp(CLng(oRecordset.Fields("StartDate").Value),oRequest("StartDate").Item,vbBinaryCompare) = 0Then
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
				End If
                If bForExport Then
                    sTextBegin ="=T("""
                    sTextEnd=""")"
                End If

				sFontBegin = ""
				sFontEnd = ""
				If Not (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And (CInt(oRecordset.Fields("StatusID").Value) = 1) Then
					sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
					sFontEnd = "</FONT>"
				End If
				If (Not bForExport) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					sRowContents = "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Positions&PositionsAction=1&PositionID=" & CLng(oRecordset.Fields("PositionID").Value) & "&StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & "&Delete=1" & sPage & sFilter & """>"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
					sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Positions&PositionsAction=1&PositionID=" & CLng(oRecordset.Fields("PositionID").Value) & "&StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & "&Change=1" & sPage & sFilter & """>"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
					sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Positions&PositionsAction=1&PositionID=" & CLng(oRecordset.Fields("PositionID").Value) & "&StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & "&Apply=1" & sPage & sFilter & """>"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
				'sRowContents = sRowContents & TABLE_SEPARATOR 
				'aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
				'sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionLongName").Value)) & sBoldEnd & sFontEnd
				If CInt(oRecordset.Fields("CompanyID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value)) & sBoldEnd & sFontEnd
				End If
                If CInt(oRecordset.Fields("EmployeeTypeID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & sBoldEnd & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
				If CInt(oRecordset.Fields("LevelID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					'sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value)) & sBoldEnd & sFontEnd
					sLevelShortName = CStr(oRecordset.Fields("LevelName").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("GroupGradeLevelID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					'sLevelShortName = CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("IntegrationID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CInt(oRecordset.Fields("IntegrationID").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("ClassificationID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CInt(oRecordset.Fields("ClassificationID").Value)) & sBoldEnd & sFontEnd
				End If
				If CSng(oRecordset.Fields("EconomicZoneID").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneName").Value)) & sBoldEnd & sFontEnd
				End If
				If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CSng(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("BranchID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BranchName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("SubBranchID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("SubBranchName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("HierarchyID").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("HierarchyID").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("GenericPositionID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("No Aplica") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("GenericPositionName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("AuthorizedJobs").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("NA") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CInt(oRecordset.Fields("AuthorizedJobs").Value)) & sBoldEnd & sFontEnd
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & sBoldEnd & sFontEnd
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & sBoldEnd & sFontEnd
				End If
				If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And (CInt(oRecordset.Fields("StatusID").Value) = 1) Then
					sFontBegin = ""
					sFontEnd = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
					sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&StartPage=" & CInt(oRequest("StartPage").Item) & "&RecordID=" & CLng(oRecordset.Fields("RecordID").Value) & "&ConceptID=" & CLng(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
				Else
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd
				End If
                If bForExport Then
                    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value)) & sBoldEnd & sFontEnd
                End If
				sRowContents = sRowContents & TABLE_SEPARATOR
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
			Response.Write "</TABLE></DIV><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen tabuladores de pago registrados en el sistema para el concepto seleccionado."
		End If
	End If

	Set oRecordset = Nothing
	DisplayPositionsTable = lErrorNumber
	Err.Clear
End Function

Function GetPositionn(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an area from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionn"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePositionComponent(oRequest, aPositionComponent)
	End If

	If ((aPositionComponent(N_ID_POSITION) = -1) Or (aPositionComponent(N_START_DATE_POSITION) = 0)) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto y su fecha de inicio para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate = " & aPositionComponent(N_START_DATE_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aPositionComponent(S_SHORT_NAME_POSITION) = CStr(oRecordset.Fields("PositionShortName").Value)
				aPositionComponent(S_NAME_POSITION) = CStr(oRecordset.Fields("PositionName").Value)
				aPositionComponent(S_LONG_NAME_POSITION) = CStr(oRecordset.Fields("PositionLongName").Value)
				aPositionComponent(S_DESCRIPTION_POSITION) = CStr(oRecordset.Fields("PositionDescription").Value)
				aPositionComponent(N_START_DATE_POSITION) = CLng(oRecordset.Fields("StartDate").Value)
				aPositionComponent(N_END_DATE_POSITION) = CLng(oRecordset.Fields("EndDate").Value)
				aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				aPositionComponent(N_POSITION_TYPE_ID_POSITION) = CLng(oRecordset.Fields("PositionTypeID").Value)
				aPositionComponent(N_COMPANY_ID_POSITION) = CLng(oRecordset.Fields("CompanyID").Value)
				aPositionComponent(N_CLASSIFICATION_ID_POSITION) = CLng(oRecordset.Fields("ClassificationID").Value)
				aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
				aPositionComponent(N_INTEGRATION_ID_POSITION) = CLng(oRecordset.Fields("IntegrationID").Value)
				aPositionComponent(N_LEVEL_ID_POSITION) = CLng(oRecordset.Fields("LevelID").Value)
				aPositionComponent(N_BRANCH_ID_POSITION) = CLng(oRecordset.Fields("BranchID").Value)
				aPositionComponent(N_SUB_BRANCH_ID_POSITION) = CLng(oRecordset.Fields("SubBranchID").Value)
				aPositionComponent(N_HIERARCHY_ID_POSITION) = CLng(oRecordset.Fields("HierarchyID").Value)
				aPositionComponent(N_GENERIC_POSITION_ID_POSITION) = CLng(oRecordset.Fields("GenericPositionID").Value)
				aPositionComponent(D_WORKING_HOURS_POSITION) = CDbl(oRecordset.Fields("WorkingHours").Value)
				aPositionComponent(N_STRATEGIC_POSITION) = CLng(oRecordset.Fields("Strategic").Value)
				aPositionComponent(N_NOMINATION_POSITION) = CLng(oRecordset.Fields("Nomination").Value)
				aPositionComponent(N_STATUS_ID_POSITION) = CLng(oRecordset.Fields("StatusID").Value)
				aPositionComponent(N_ACTIVE_POSITION) = CLng(oRecordset.Fields("Active").Value)
				aPositionComponent(N_DEPRECIATED_POSITION) = CLng(oRecordset.Fields("Depreciated").Value)
				aPositionComponent(N_ECONOMICZONE) = CLng(oRecordset.Fields("EconomicZoneID").value) 
				aPositionComponent(S_COMMENTS) = CStr(oRecordset.Fields("Comments").value)
                aPositionComponent(D_AUTHORIZED_JOBS) = CLng(oRecordset.Fields("AuthorizedJobs").value)
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetPositionn = lErrorNumber
	Err.Clear
End Function

Function GetPositions(oRequest, oADODBConnection, aPositionComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the areas from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositions"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePositionComponent(oRequest, aPositionComponent)
	End If

    If Len(oRequest("ApplyFilter").Item) > 0 Then
		Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "Positions.StartDate", False, sCondition)
		aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & sCondition
		If CInt(oRequest("PositionShortNameFilter").Item) > 0 Then
			aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & " And (Positions.PositionShortName Like '" & S_WILD_CHAR & oRequest("PositionShortNameFilter").Item & S_WILD_CHAR & "')"
		End If
		If CInt(oRequest("GroupGradeLevelIDFilter").Item) > 0 Then
			aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & " And (Positions.GroupGradeLevelID = " & oRequest("GroupGradeLevelIDFilter").Item & ")"
		End If
        If CInt(oRequest("EmployeeTypeIDFilter").Item) >= 0 Then
            aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & " And (Positions.EmployeeTypeID = " & oRequest("EmployeeTypeIDFilter").Item & ")"
        End If
	End If

	If (Len(aPositionComponent(S_QUERY_CONDITION_POSITION)) > 0) Then
		sCondition = Trim(aPositionComponent(S_QUERY_CONDITION_POSITION))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.*, EmployeeTypes.EmployeeTypeShortName, EmployeeTypeName, PositionTypes.PositionTypeShortName, PositionTypeName, Companies.CompanyShortName, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, GroupGradeLevelName, Levels.LevelName, Branches.BranchShortName, BranchName, SubBranches.SubBranchShortName, SubBranchName, GenericPositions.GenericPositionName, EconomicZones.EconomicZoneName From Positions, EmployeeTypes, PositionTypes, Companies, GroupGradeLevels, Levels, Branches, SubBranches, GenericPositions,EconomicZones Where (Positions.PositionID > -1) And (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.BranchID=Branches.BranchID) And (Positions.SubBranchID=SubBranches.SubBranchID) And (Positions.GenericPositionID=GenericPositions.GenericPositionID) And(Positions.EconomicZoneID=EconomicZones.EconomicZoneID)  "& sCondition &" Order By PositionShortName, PositionName, PositionID, Positions.StartDate;", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetPositions = lErrorNumber
	Err.Clear
End Function

Function GetPositions1(oRequest, oADODBConnection, aPositionComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositions1"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sCondition

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aPositionComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los puestos."
	If Len(aPositionComponent(S_QUERY_CONDITION_CONCEPT)) > 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where " & aPositionComponent(S_QUERY_CONDITION_CONCEPT), "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID)" & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID"" />"

	GetPositions1 = lErrorNumber
	Err.Clear
End Function

Function GetPositionCrossType(oADODBConnection, aPositionComponent, sPositionCrossType, lPositionID, lPositionStartDate, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing for the
'         record to insert
'Inputs:  oADODBConnection, aPositionComponent
'Outputs: sPositionCrossType, lPositionID, lPositionStartDate, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionCrossType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	sQuery = "Select * From Positions Where (PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And (StartDate<>" & aPositionComponent(N_START_DATE_POSITION) & ")" & _
            " And (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "')" & _
			" And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ")" & _
			" And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ")" & _
			" And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ")" & _
			" And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ")" & _
			" And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ")" & _
			" And (StartDate<" & aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate>" & aPositionComponent(N_END_DATE_POSITION) & ") Order By StartDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sPositionCrossType = "Cross"
		Else
	        sQuery = "Select * From Positions Where (PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And (StartDate<>" & aPositionComponent(N_START_DATE_POSITION) & ")" & _
                    " And (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "')" & _
					" And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ")" & _
					" And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ")" & _
					" And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ")" & _
					" And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ")" & _
					" And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ")" & _
					" And (StartDate>" & aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate<" & aPositionComponent(N_END_DATE_POSITION) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sPositionCrossType = "Inner"
				Else
	                sQuery = "Select * From Positions Where (PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And (StartDate<>" & aPositionComponent(N_START_DATE_POSITION) & ")" & _
                            " And (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "')" & _
							" And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ")" & _
							" And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ")" & _
							" And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ")" & _
							" And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ")" & _
							" And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ")" & _
							" And (StartDate<" & aPositionComponent(N_START_DATE_POSITION) & ") And ((EndDate<=" & aPositionComponent(N_END_DATE_POSITION) & ") And (EndDate>=" & aPositionComponent(N_START_DATE_POSITION) & ")) And (StartDate<EndDate) Order By StartDate Desc"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sPositionCrossType = "Left"
                            lPositionID = CLng(oRecordset.Fields("PositionID").Value)
                            lPositionStartDate = CLng(oRecordset.Fields("StartDate").Value)
						Else
	                        sQuery = "Select * From Positions Where (PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And (StartDate<>" & aPositionComponent(N_START_DATE_POSITION) & ")" & _
                                    " And (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "')" & _
									" And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ")" & _
									" And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ")" & _
									" And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ")" & _
									" And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ")" & _
									" And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ")" & _
									" And (StartDate>=" & aPositionComponent(N_START_DATE_POSITION) & ") And ((EndDate>=" & aPositionComponent(N_END_DATE_POSITION) & ") And (StartDate<=" & aPositionComponent(N_END_DATE_POSITION) & ")) And (StartDate<EndDate) Order By StartDate"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sPositionCrossType = "Right"
								End If
							Else
								lErrorNumber = -1
								sErrorDescription = "No se pudo verifiar si el registro del puesto se empalma con otros puestos ya registrados."
							End If
						End If
					Else
						lErrorNumber = -1
						sErrorDescription = "No se pudo verifiar si el registro del puesto se empalma con otros puestos ya registrados."
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No se pudo verifiar si el registro del puesto se empalma con otros puestos ya registrados."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verifiar si el registro del puesto se empalma con otros puestos ya registrados."
	End If

	Set oRecordset = Nothing
	GetPositionCrossType = lErrorNumber
	Err.Clear
End Function

Function PositionHasChanged(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PositionHasChanged"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aPositionComponent)
	End If

	If aPositionComponent(N_ID_POSITION) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate = " & aPositionComponent(N_START_DATE_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				If (aPositionComponent(N_ID_POSITION) = CLng(oRecordset.Fields("PositionID").Value)) And _
					(aPositionComponent(S_SHORT_NAME_POSITION) = CStr(oRecordset.Fields("PositionShortName").Value)) And _
                    (aPositionComponent(S_NAME_POSITION) = CStr(oRecordset.Fields("PositionName").Value)) And _
					(aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) = CInt(oRecordset.Fields("EmployeeTypeID").Value)) And _
					(aPositionComponent(N_POSITION_TYPE_ID_POSITION) = CInt(oRecordset.Fields("PositionTypeID").Value)) And _
					(aPositionComponent(N_COMPANY_ID_POSITION) = CInt(oRecordset.Fields("CompanyID").Value)) And _
					(aPositionComponent(N_CLASSIFICATION_ID_POSITION) = CInt(oRecordset.Fields("ClassificationID").Value)) And _
					(aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) = CInt(oRecordset.Fields("GroupGradeLevelID").Value)) And _
					(aPositionComponent(N_INTEGRATION_ID_POSITION) = CInt(oRecordset.Fields("IntegrationID").Value)) And _
					(aPositionComponent(N_LEVEL_ID_POSITION) = CInt(oRecordset.Fields("LevelID").Value)) And _
					(aPositionComponent(N_BRANCH_ID_POSITION) = CInt(oRecordset.Fields("BranchID").Value)) And _
					(aPositionComponent(N_SUB_BRANCH_ID_POSITION) = CInt(oRecordset.Fields("SubBranchID").Value)) And _
					(aPositionComponent(N_HIERARCHY_ID_POSITION) = CInt(oRecordset.Fields("HierarchyID").Value)) And _
					(aPositionComponent(N_GENERIC_POSITION_ID_POSITION) = CInt(oRecordset.Fields("GenericPositionID").Value)) And _
					(aPositionComponent(D_WORKING_HOURS_POSITION) = CDbl(oRecordset.Fields("WorkingHours").Value)) And _
					(aPositionComponent(N_STRATEGIC_POSITION) = CInt(oRecordset.Fields("Strategic").Value)) And _
					(aPositionComponent(N_NOMINATION_POSITION) = CInt(oRecordset.Fields("Nomination").Value)) And _
					(aPositionComponent(N_ECONOMICZONE) = CLng(oRecordset.Fields("EconomicZoneID").Value)) _
				Then
					PositionHasChanged = False
				Else
					PositionHasChanged = True
				End If
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function ModifyPosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPosition"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
    Dim lPositionID
    Dim lPositionStartDate

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePositionComponent(oRequest, aPositionComponent)
	End If

	If Not CheckPositionInformationConsistency(aPositionComponent, sErrorDescription) Then
		lErrorNumber = -1
	Else
        lPositionStartDate = aPositionComponent(N_START_DATE_POSITION)
		If Not PositionHasChanged(oRequest, oADODBConnection, aPositionComponent, sErrorDescription) Then
			sErrorDescription = "No se pudo modificar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set AuthorizedJobs=" & aPositionComponent(D_AUTHORIZED_JOBS) & ", Comments='" & Replace(aPositionComponent(S_COMMENTS), "'", "´") & " Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate=" & aPositionComponent(N_START_DATE_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
            If aPositionComponent(N_APPLIED_DATE_POSITION) = -1 Then
                lErrorNumber = -1
                sErrorDescription = "No se pudo modificar la información del puesto debido a que no se indico la quincena desde la que se aplicara el cambio."
            Else
                aPositionComponent(N_START_DATE_POSITION) = aPositionComponent(N_APPLIED_DATE_POSITION)
			    If Not CheckExistencyOfPosition(aPositionComponent, lPositionID, lPositionStartDate, sErrorDescription) Then
				    lErrorNumber = L_ERR_DUPLICATED_RECORD
				    'sErrorDescription = "No se pude modificar la información del puesto, debido a que existen registros con las mismas condiciones en el periodo indicado."
			    Else
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set EndDate = " & AddDaysToSerialDate(aPositionComponent(N_START_DATE_POSITION), -1) & " Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate=" & lPositionStartDate & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				    'sErrorDescription = "No se pudo modificar la información del puesto."
				    'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set PositionShortName='" & Replace(aPositionComponent(S_SHORT_NAME_POSITION), "'", "") & "', PositionName='" & Replace(aPositionComponent(S_NAME_POSITION), "'", "´") & "', PositionLongName='" & Replace(aPositionComponent(S_LONG_NAME_POSITION), "'", "´") & "', PositionDescription='" & Replace(aPositionComponent(S_DESCRIPTION_POSITION), "'", "´") & "', StartDate=" & aPositionComponent(N_START_DATE_POSITION) & ", EndDate=" & aPositionComponent(N_END_DATE_POSITION) & ", EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ", PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ", CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ", ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ", GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ", IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ", LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ", BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ", SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ", HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ", GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ", WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ", Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ", Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ", StatusID=" & aPositionComponent(N_STATUS_ID_POSITION) & ", Active=" & aPositionComponent(N_ACTIVE_POSITION) & ", Depreciated=" & aPositionComponent(N_DEPRECIATED_POSITION) & ",EconomicZoneID=" & aPositionComponent(N_ECONOMICZONE) & ", Comments='" & Replace(aPositionComponent(S_COMMENTS), "'", "´") & "', AuthorizedJobs=" & aPositionComponent(D_AUTHORIZED_JOBS) & " Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate=" & aPositionComponent(N_START_DATE_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
                    'aPositionComponent(N_START_DATE_POSITION) = lPositionStartDate
                    lErrorNumber = AddPosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
			    End If
            End If
		End If
	End If

	ModifyPosition = lErrorNumber
	Err.Clear
End Function

Function ModifyPosition1(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPosition1"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sSpecialCondition
	Dim bHasChanged
	Dim lNewRecordID

	If aPositionComponent(N_ID_POSITION) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aPositionComponent(N_END_DATE_POSITION) = 0 Then aPositionComponent(N_END_DATE_POSITION) = 30000000
		bHasChanged = True
		If Not PositionHasChanged(oRequest, oADODBConnection, aPositionComponent, sErrorDescription) Then
			bHasChanged = False
			sSpecialCondition = "(PositionID<>" & aPositionComponent(N_ID_POSITION) & ") And"
		End If
		sQuery = "Select * From Positions Where " & sSpecialCondition & " (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "') And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ") And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ") And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ") And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ") And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ") And (StartDate<" & aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate>" & aPositionComponent(N_END_DATE_POSITION) & ") Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set EndDate=" & AddDaysToSerialDate(aPositionComponent(N_START_DATE_POSITION), -1) & " Where (PositionID=" & oRecordset.Fields("PositionID").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "Positions", "PositionID", "", 1, lNewRecordID, sErrorDescription)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Positions (PositionID, PositionShortName, PositionName, PositionLongName, PositionDescription, StartDate, EndDate, EmployeeTypeID, PositionTypeID, CompanyID, ClassificationID, GroupGradeLevelID, IntegrationID, LevelID, BranchID, SubBranchID, HierarchyID, GenericPositionID, WorkingHours, Strategic, Nomination, StatusID, Active, Depreciated) Values (" & lNewRecordID & ", '" & oRecordset.Fields("PositionShortName").Value & "', '" & oRecordset.Fields("PositionName").Valuea & "', '" & oRecordset.Fields("PositionLongName").Value & "', '" & oRecordset.Fields("PositionDescription").Value & "', " & AddDaysToSerialDate(aPositionComponent(N_END_DATE_POSITION), 1) & ", " & oRecordset.Fields("EndDate").Value & ", " & oRecordset.Fields("EmployeeTypeID").Value & ", " & oRecordset.Fields("PositionTypeID").Value & ", " & oRecordset.Fields("CompanyID").Value & ", " & oRecordset.Fields("ClassificationID").Value & ", " & oRecordset.Fields("GroupGradeLevelID").Value & ", " & oRecordset.Fields("IntegrationID").Value & ", " & oRecordset.Fields("LevelID").Value & ", " & oRecordset.Fields("BranchID").Value & ", " & oRecordset.Fields("SubBranchID").Value & ", " & oRecordset.Fields("HierarchyID").Value & ", " & oRecordset.Fields("GenericPositionID").Value & ", " & oRecordset.Fields("WorkingHours").Value & ", " & oRecordset.Fields("Strategic").Value & ", " & oRecordset.Fields("Nomination").Value & ", " & oRecordset.Fields("StatusID").Value & ", " & oRecordset.Fields("Active").Value & ", " & oRecordset.Fields("Depreciated").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sQuery = "Select * From Positions Where " & sSpecialCondition & " (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "') And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ") And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ") And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ") And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ") And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ") And (StartDate>=" & aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate<=" & aPositionComponent(N_END_DATE_POSITION) & ") Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Positions Where (PositionID=" & oRecordset.Fields("PositionID").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
		End If
		sQuery = "Select * From Positions Where " & sSpecialCondition & " (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "') And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ") And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ") And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ") And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ") And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ") And (StartDate<" & aPositionComponent(N_START_DATE_POSITION) & ") And ((EndDate<=" & aPositionComponent(N_END_DATE_POSITION) & ") And (EndDate>=" & aPositionComponent(N_START_DATE_POSITION) & ")) And (StartDate<EndDate) Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set EndDate=" & AddDaysToSerialDate(aPositionComponent(N_START_DATE_POSITION), -1) & " Where (PositionID=" & oRecordset.Fields("PositionID").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sQuery = "Select * From Positions Where " & sSpecialCondition & " (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "') And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ") And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ") And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ") And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ") And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ") And (StartDate>=" & aPositionComponent(N_START_DATE_POSITION) & ") And ((EndDate>" & aPositionComponent(N_END_DATE_POSITION) & ") And (StartDate<=" & aPositionComponent(N_START_DATE_POSITION) & ")) And (StartDate<EndDate) Order By StartDate"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set StartDate=" & AddDaysToSerialDate(aPositionComponent(N_END_DATE_POSITION), 1) & " Where (PositionID=" & oRecordset.Fields("PositionID").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sErrorDescription = "No se pudo guardar la información del nuevo registro."
		If Not bHasChanged Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "', PositionName='" & aPositionComponent(S_NAME_POSITION) & "', PositionLongName='" & aPositionComponent(S_LONG_NAME_POSITION) & "', PositionDescription='" & aPositionComponent(S_DESCRIPTION_POSITION) & "', StartDate=" & aPositionComponent(N_START_DATE_POSITION) & ", EndDate=" & aPositionComponent(N_END_DATE_POSITION) & ", StatusID=" & aPositionComponent(N_STATUS_ID_POSITION) & ", Active=" & aPositionComponent(N_ACTIVE_POSITION) & ", Depreciated=" & aPositionComponent(N_DEPRECIATED_POSITION) & "  Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Positions", "PositionID", "", 1, lNewRecordID, sErrorDescription)
			If lErrorNumber = 0 Then
				sQuery = "Select * From Positions Where " & sSpecialCondition & " (PositionShortName='" & aPositionComponent(S_SHORT_NAME_POSITION) & "') And (EmployeeTypeID=" & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ") And (PositionTypeID=" & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ") And (CompanyID=" & aPositionComponent(N_COMPANY_ID_POSITION) & ") And (ClassificationID=" & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ") And (GroupGradeLevelID=" & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ") And (IntegrationID=" & aPositionComponent(N_INTEGRATION_ID_POSITION) & ") And (LevelID=" & aPositionComponent(N_LEVEL_ID_POSITION) & ") And (BranchID=" & aPositionComponent(N_BRANCH_ID_POSITION) & ") And (SubBranchID=" & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ") And (HierarchyID=" & aPositionComponent(N_HIERARCHY_ID_POSITION) & ") And (GenericPositionID=" & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ") And (WorkingHours=" & aPositionComponent(D_WORKING_HOURS_POSITION) & ") And (Strategic=" & aPositionComponent(N_STRATEGIC_POSITION) & ") And (Nomination=" & aPositionComponent(N_NOMINATION_POSITION) & ") And (StartDate=" & aPositionComponent(N_START_DATE_POSITION) & ") And (EndDate=" & aPositionComponent(N_END_DATE_POSITION) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Positions Set StartDate=" & AddDaysToSerialDate(aPositionComponent(N_END_DATE_POSITION), 1) & " Where (PositionID=" & oRecordset.Fields("PositionID").Value & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						aPositionComponent(N_STATUS_ID_POSITION) = 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Positions (PositionID, PositionShortName, PositionName, PositionLongName, PositionDescription, StartDate, EndDate, EmployeeTypeID, PositionTypeID, CompanyID, ClassificationID, GroupGradeLevelID, IntegrationID, LevelID, BranchID, SubBranchID, HierarchyID, GenericPositionID, WorkingHours, Strategic, Nomination, StatusID, Active, Depreciated) Values (" & lNewRecordID & ", '" & Replace(aPositionComponent(S_SHORT_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_LONG_NAME_POSITION), "'", "") & "', '" & Replace(aPositionComponent(S_DESCRIPTION_POSITION), "'", "´") & "', " & aPositionComponent(N_START_DATE_POSITION) & ", " & aPositionComponent(N_END_DATE_POSITION) & ", " & aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) & ", " & aPositionComponent(N_POSITION_TYPE_ID_POSITION) & ", " & aPositionComponent(N_COMPANY_ID_POSITION) & ", " & aPositionComponent(N_CLASSIFICATION_ID_POSITION) & ", " & aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) & ", " & aPositionComponent(N_INTEGRATION_ID_POSITION) & ", " & aPositionComponent(N_LEVEL_ID_POSITION) & ", " & aPositionComponent(N_BRANCH_ID_POSITION) & ", " & aPositionComponent(N_SUB_BRANCH_ID_POSITION) & ", " & aPositionComponent(N_HIERARCHY_ID_POSITION) & ", " & aPositionComponent(N_GENERIC_POSITION_ID_POSITION) & ", " & aPositionComponent(D_WORKING_HOURS_POSITION) & ", " & aPositionComponent(N_STRATEGIC_POSITION) & ", " & aPositionComponent(N_NOMINATION_POSITION) & ", " & aPositionComponent(N_STATUS_ID_POSITION) & ", " & aPositionComponent(N_ACTIVE_POSITION) & ", " & aPositionComponent(N_DEPRECIATED_POSITION) & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	ModifyPosition1 = lErrorNumber
	Err.Clear
End Function

Function RemovePosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPositionComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePosition"
	Dim lErrorNumber
	Dim bComponentInitialized
    Dim sQuery

	bComponentInitialized = aPositionComponent(B_POSITION_COMPONENT_INITIALIZED)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aPositionComponent)
	End If

	If ((aPositionComponent(N_ID_POSITION) = -1) Or (aPositionComponent(N_START_DATE_POSITION) = 0)) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto y su fecha de inicio para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sQuery = "Select * From ConceptsValues Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
                lErrorNumber = -1
                sErrorDescription = "No se puede eliminar el puesto debido a que ya se registro un tabulador de pago para este."
            Else
		        sQuery = "Select * From Jobs Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ")"
		        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		        If lErrorNumber = 0 Then
			        If Not oRecordset.EOF Then
                        lErrorNumber = -1
                        sErrorDescription = "No se puede eliminar el puesto debido a que ya se registro una paza que utiliza este puesto."
                    Else
		                sErrorDescription = "No se pudo eliminar la información del puesto."
		                lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Positions Where (PositionID=" & aPositionComponent(N_ID_POSITION) & ") And (StartDate = " & oRequest("StartDate").Item & ")", "PositionComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
                    End If
                Else
                    sErrorDescription = "No se puede eliminar el puesto debido a que no se pudo validar si ya fue registrada una paza que utiliza este puesto."
                End If
            End If
        Else
            sErrorDescription = "No se puede eliminar el puesto debido a que no se pudo validar si ya fue registrado un tabulador de pago para este."
        End If
	End If

	RemovePosition = lErrorNumber
	Err.Clear
End Function

Function CheckPositionInformationConsistency(aPositionComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aPositionComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckPositionInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aPositionComponent(N_ID_POSITION)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del puesto no es un valor numérico."
		bIsCorrect = False
	End If
	If StrComp(oRequest("Action").Item, "Positions", vbBinaryCompare) = 0 Then
		If Not IsNumeric(aPositionComponent(N_START_DATE_POSITION)) Then aPositionComponent(N_START_DATE_POSITION) = Left(GetSerialNumberForDate(""), Len("00000000"))
		If Not IsNumeric(aPositionComponent(N_END_DATE_POSITION)) Then aPositionComponent(N_END_DATE_POSITION) = 30000000
		If Len(aPositionComponent(S_SHORT_NAME_POSITION)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave del puesto está vacía."
			bIsCorrect = False
		End If
		If Len(aPositionComponent(S_NAME_POSITION)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del puesto está vacío."
			bIsCorrect = False
		End If
		If Len(aPositionComponent(S_LONG_NAME_POSITION)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre largo del puesto está vacío."
			bIsCorrect = False
		End If
		
		If Not IsNumeric(aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION)) Then aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_POSITION_TYPE_ID_POSITION)) Then aPositionComponent(N_POSITION_TYPE_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_COMPANY_ID_POSITION)) Then aPositionComponent(N_COMPANY_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_CLASSIFICATION_ID_POSITION)) Then aPositionComponent(N_CLASSIFICATION_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION)) Then aPositionComponent(N_GROUP_GRADE_LEVEL_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_INTEGRATION_ID_POSITION)) Then aPositionComponent(N_INTEGRATION_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_LEVEL_ID_POSITION)) Then aPositionComponent(N_LEVEL_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_BRANCH_ID_POSITION)) Then aPositionComponent(N_BRANCH_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_SUB_BRANCH_ID_POSITION)) Then aPositionComponent(N_SUB_BRANCH_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_HIERARCHY_ID_POSITION)) Then aPositionComponent(N_HIERARCHY_ID_POSITION) = 0
		If Not IsNumeric(aPositionComponent(N_GENERIC_POSITION_ID_POSITION)) Then aPositionComponent(N_GENERIC_POSITION_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(D_WORKING_HOURS_POSITION)) Then aPositionComponent(D_WORKING_HOURS_POSITION) = 0
		If Not IsNumeric(aPositionComponent(N_STRATEGIC_POSITION)) Then aPositionComponent(N_STRATEGIC_POSITION) = 1
		If Not IsNumeric(aPositionComponent(N_NOMINATION_POSITION)) Then aPositionComponent(N_NOMINATION_POSITION) = 1
		If Not IsNumeric(aPositionComponent(N_STATUS_ID_POSITION)) Then aPositionComponent(N_STATUS_ID_POSITION) = -1
		If Not IsNumeric(aPositionComponent(N_ACTIVE_POSITION)) Then aPositionComponent(N_ACTIVE_POSITION) = 1
		If Not IsNumeric(aPositionComponent(N_DEPRECIATED_POSITION)) Then aPositionComponent(N_DEPRECIATED_POSITION) = 0
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PositionComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckPositionInformationConsistency = bIsCorrect
	Err.Clear
End Function
%>