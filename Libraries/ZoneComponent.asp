<%
Const N_ID_ZONE = 0
Const N_PARENT_ID_ZONE = 1
Const N_ZONE_TYPE_ID_ZONE = 2
Const S_CODE_ZONE = 3
Const S_NAME_ZONE = 4
Const S_DESCRIPTION_ZONE = 5
Const S_PATH_ZONE = 6
Const N_START_DATE_ZONE = 7
Const N_END_DATE_ZONE = 8
Const N_ACTIVE_ZONE = 9
Const S_QUERY_CONDITION_ZONE = 10
Const B_CHECK_FOR_DUPLICATED_ZONE = 11
Const B_IS_DUPLICATED_ZONE = 12
Const B_COMPONENT_INITIALIZED_ZONE = 13

Const N_ZONE_COMPONENT_SIZE = 13

Dim aZoneComponent()
Redim aZoneComponent(N_ZONE_COMPONENT_SIZE)

Function InitializeZoneComponent(oRequest, aZoneComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Zone
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aZoneComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeZoneComponent"
	Redim Preserve aZoneComponent(N_ZONE_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aZoneComponent(N_ID_ZONE)) Then
		If Len(oRequest("ZoneID").Item) > 0 Then
			aZoneComponent(N_ID_ZONE) = CLng(oRequest("ZoneID").Item)
		Else
			aZoneComponent(N_ID_ZONE) = -1
		End If
	End If

	If IsEmpty(aZoneComponent(N_PARENT_ID_ZONE)) Then
		If Len(oRequest("ParentID").Item) > 0 Then
			aZoneComponent(N_PARENT_ID_ZONE) = CLng(oRequest("ParentID").Item)
		Else
			aZoneComponent(N_PARENT_ID_ZONE) = -1
		End If
	End If

	If IsEmpty(aZoneComponent(N_ZONE_TYPE_ID_ZONE)) Then
		If Len(oRequest("ZoneTypeID").Item) > 0 Then
			aZoneComponent(N_ZONE_TYPE_ID_ZONE) = CLng(oRequest("ZoneTypeID").Item)
		Else
			aZoneComponent(N_ZONE_TYPE_ID_ZONE) = 1
		End If
	End If

	If IsEmpty(aZoneComponent(S_CODE_ZONE)) Then
		If Len(oRequest("ZoneCode").Item) > 0 Then
			aZoneComponent(S_CODE_ZONE) = oRequest("ZoneCode").Item
		Else
			aZoneComponent(S_CODE_ZONE) = ""
		End If
	End If
	aZoneComponent(S_CODE_ZONE) = Left(aZoneComponent(S_CODE_ZONE), 5)

	If IsEmpty(aZoneComponent(S_NAME_ZONE)) Then
		If Len(oRequest("ZoneName").Item) > 0 Then
			aZoneComponent(S_NAME_ZONE) = oRequest("ZoneName").Item
		Else
			aZoneComponent(S_NAME_ZONE) = ""
		End If
	End If
	aZoneComponent(S_NAME_ZONE) = Left(aZoneComponent(S_NAME_ZONE), 255)

	If IsEmpty(aZoneComponent(S_DESCRIPTION_ZONE)) Then
		If Len(oRequest("Description").Item) > 0 Then
			aZoneComponent(S_DESCRIPTION_ZONE) = oRequest("Description").Item
		Else
			aZoneComponent(S_DESCRIPTION_ZONE) = ""
		End If
	End If
	aZoneComponent(S_DESCRIPTION_ZONE) = Left(aZoneComponent(S_DESCRIPTION_ZONE), 2000)

	If IsEmpty(aZoneComponent(S_PATH_ZONE)) Then
		If Len(oRequest("ZonePath").Item) > 0 Then
			aZoneComponent(S_PATH_ZONE) = oRequest("ZonePath").Item
		Else
			aZoneComponent(S_PATH_ZONE) = ",-1,"
			If aZoneComponent(N_ID_ZONE) > -1 Then aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & aZoneComponent(N_ID_ZONE) & ","
		End If
	End If
	aZoneComponent(S_PATH_ZONE) = Left(aZoneComponent(S_PATH_ZONE), 255)

	If IsEmpty(aZoneComponent(N_START_DATE_ZONE)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aZoneComponent(N_START_DATE_ZONE) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aZoneComponent(N_START_DATE_ZONE) = CLng(oRequest("StartDate").Item)
		Else
			aZoneComponent(N_START_DATE_ZONE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aZoneComponent(N_END_DATE_ZONE)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aZoneComponent(N_END_DATE_ZONE) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aZoneComponent(N_END_DATE_ZONE) = CLng(oRequest("EndDate").Item)
		Else
			aZoneComponent(N_END_DATE_ZONE) = 30000000
		End If
	End If

	If IsEmpty(aZoneComponent(N_ACTIVE_ZONE)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aZoneComponent(N_ACTIVE_ZONE) = CInt(oRequest("Active").Item)
		Else
			aZoneComponent(N_ACTIVE_ZONE) = 1
		End If
	End If

	aZoneComponent(S_QUERY_CONDITION_ZONE) = ""
	aZoneComponent(B_CHECK_FOR_DUPLICATED_ZONE) = True
	aZoneComponent(B_IS_DUPLICATED_ZONE) = False

	aZoneComponent(B_COMPONENT_INITIALIZED_ZONE) = True
	InitializeZoneComponent = Err.number
	Err.Clear
End Function

Function AddZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new zone into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddZone"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_ID_ZONE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Zones", "ZoneID", "", 1, aZoneComponent(N_ID_ZONE), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aZoneComponent(B_CHECK_FOR_DUPLICATED_ZONE) Then
			lErrorNumber = CheckExistencyOfZone(aZoneComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aZoneComponent(B_IS_DUPLICATED_ZONE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe una zona con el código " & aZoneComponent(S_CODE_ZONE) & " o el nombre " & aZoneComponent(S_NAME_ZONE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckZoneInformationConsistency(aZoneComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Zones (ZoneID, ParentID, ZoneTypeID, ZoneCode, ZoneName, Description, ZonePath, StartDate, EndDate, ModifyDate, UserID, Active) Values (" & aZoneComponent(N_ID_ZONE) & ", " & aZoneComponent(N_PARENT_ID_ZONE) & ", " & aZoneComponent(N_ZONE_TYPE_ID_ZONE) & ", '" & Replace(aZoneComponent(S_CODE_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_NAME_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_DESCRIPTION_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_PATH_ZONE), "'", "") & "', " & aZoneComponent(N_START_DATE_ZONE) & ", " & aZoneComponent(N_END_DATE_ZONE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aZoneComponent(N_ACTIVE_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	AddZone = lErrorNumber
	Err.Clear
End Function

Function GetZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a zone from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetZone"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_ID_ZONE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Zones Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ") And (StartDate=" & aZoneComponent(N_START_DATE_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aZoneComponent(N_PARENT_ID_ZONE) = CLng(oRecordset.Fields("ParentID").Value)
				aZoneComponent(N_ZONE_TYPE_ID_ZONE) = CLng(oRecordset.Fields("ZoneTypeID").Value)
				aZoneComponent(S_CODE_ZONE) = CStr(oRecordset.Fields("ZoneCode").Value)
				aZoneComponent(S_NAME_ZONE) = CStr(oRecordset.Fields("ZoneName").Value)
				aZoneComponent(S_DESCRIPTION_ZONE) = CStr(oRecordset.Fields("Description").Value)
				aZoneComponent(S_PATH_ZONE) = CStr(oRecordset.Fields("ZonePath").Value)
				aZoneComponent(N_START_DATE_ZONE) = CLng(oRecordset.Fields("StartDate").Value)
				aZoneComponent(N_END_DATE_ZONE) = CLng(oRecordset.Fields("EndDate").Value)
				aZoneComponent(N_ACTIVE_ZONE) = CInt(oRecordset.Fields("Active").Value)
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetZone = lErrorNumber
	Err.Clear
End Function

Function GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for a zone from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetZonePath"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_PARENT_ID_ZONE) = -1 Then
		aZoneComponent(S_PATH_ZONE) = ",-1,"
	Else
		sErrorDescription = "No se pudo obtener la ruta de la zona."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZonePath From Zones Where ZoneID=" & aZoneComponent(N_PARENT_ID_ZONE), "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aZoneComponent(S_PATH_ZONE) = ",-1,"
			Else
				aZoneComponent(S_PATH_ZONE) = CStr(oRecordset.Fields("ZonePath").Value)
			End If
			aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & aZoneComponent(N_ID_ZONE) & ","
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetZonePath = lErrorNumber
	Err.Clear
End Function

Function GetZones(oRequest, oADODBConnection, aZoneComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the zones from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetZones"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If (Len(aZoneComponent(S_QUERY_CONDITION_ZONE)) > 0) Or (aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1) Then
		sCondition = Trim(aZoneComponent(S_QUERY_CONDITION_ZONE))
		If aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1 Then
			sCondition = Trim(sCondition & " And ((ZoneID=" & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & ") Or (ZonePath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & "," & S_WILD_CHAR & "')) ")
		End If
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Zones Where (ZoneID>-1) " & sCondition & " Order By ZoneCode, ZoneName", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetZones = lErrorNumber
	Err.Clear
End Function

Function ModifyZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing zone in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyZone"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_ID_ZONE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aZoneComponent(B_CHECK_FOR_DUPLICATED_ZONE) Then
			lErrorNumber = CheckExistencyOfZone(aZoneComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aZoneComponent(B_IS_DUPLICATED_ZONE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe una zona con el código " & aZoneComponent(S_CODE_ZONE) & " o el nombre " & aZoneComponent(S_NAME_ZONE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckZoneInformationConsistency(aZoneComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo modificar la información del registro."
						If aZoneComponent(N_START_DATE_ZONE) = CLng(oRequest("StartDateOld").Item) Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Zones Set ZoneCode='" & Replace(aZoneComponent(S_CODE_ZONE), "'", "") & "', ZoneName='" & Replace(aZoneComponent(S_NAME_ZONE), "'", "") & "', Description='" & Replace(aZoneComponent(S_DESCRIPTION_ZONE), "'", "") & "', ZonePath='" & Replace(aZoneComponent(S_PATH_ZONE), "'", "") & "', ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Active=" & aZoneComponent(N_ACTIVE_ZONE) & " Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ") And (EndDate=30000000)", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Zones Set EndDate=" & aZoneComponent(N_START_DATE_ZONE) & " Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ") And (EndDate=30000000)", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							sErrorDescription = "No se pudo modificar la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Zones (ZoneID, ParentID, ZoneTypeID, ZoneCode, ZoneName, Description, ZonePath, StartDate, EndDate, ModifyDate, UserID, Active) Values (" & aZoneComponent(N_ID_ZONE) & ", " & aZoneComponent(N_PARENT_ID_ZONE) & ", " & aZoneComponent(N_ZONE_TYPE_ID_ZONE) & ", '" & Replace(aZoneComponent(S_CODE_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_NAME_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_DESCRIPTION_ZONE), "'", "´") & "', '" & Replace(aZoneComponent(S_PATH_ZONE), "'", "") & "', " & aZoneComponent(N_START_DATE_ZONE) & ", 30000000, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aZoneComponent(N_ACTIVE_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyZone = lErrorNumber
	Err.Clear
End Function

Function SetActiveForZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given zone
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForZone"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_ID_ZONE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Zones Set Active=" & CInt(oRequest("SetActive").Item) & " Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForZone = lErrorNumber
	Err.Clear
End Function

Function RemoveZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a zone from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveZone"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If aZoneComponent(N_ID_ZONE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Zones Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set ZoneID=-1 Where (ZoneID=" & aZoneComponent(N_ID_ZONE) & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveZone = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfZone(aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific zone exists in the database
'Inputs:  aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfZone"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	If Len(aZoneComponent(S_NAME_ZONE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Zones Where (ZoneID<>" & aZoneComponent(N_ID_ZONE) & ") And (ParentID=" & aZoneComponent(N_PARENT_ID_ZONE) & ") And ((ZoneCode='" & Replace(aZoneComponent(S_CODE_ZONE), "'", "") & "') Or (ZoneName='" & Replace(aZoneComponent(S_NAME_ZONE), "'", "") & "'))", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aZoneComponent(B_IS_DUPLICATED_ZONE) = True
				aZoneComponent(N_ID_ZONE) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfZone = lErrorNumber
	Err.Clear
End Function

Function CheckZoneInformationConsistency(aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aZoneComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckZoneInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aZoneComponent(N_ID_ZONE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aZoneComponent(N_PARENT_ID_ZONE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de la zona a la que pertenece esta zona no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aZoneComponent(N_ZONE_TYPE_ID_ZONE)) Then aZoneComponent(N_ZONE_TYPE_ID_ZONE) = 1
	If Len(aZoneComponent(S_CODE_ZONE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El código del registro está vacío."
		bIsCorrect = False
	End If
	If Len(aZoneComponent(S_NAME_ZONE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del registro está vacío."
		bIsCorrect = False
	End If
	If Len(aZoneComponent(S_PATH_ZONE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La ruta de la zona está vacía."
		bIsCorrect = False
	End If
	If Not IsNumeric(aZoneComponent(N_START_DATE_ZONE)) Then aZoneComponent(N_START_DATE_ZONE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Not IsNumeric(aZoneComponent(N_END_DATE_ZONE)) Then aZoneComponent(N_END_DATE_ZONE) = 30000000
	If Not IsNumeric(aZoneComponent(N_ACTIVE_ZONE)) Then aZoneComponent(N_ACTIVE_ZONE) = 1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZoneComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckZoneInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayZoneForm(oRequest, oADODBConnection, sAction, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a zone from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayZoneForm"
	Dim lErrorNumber

	If aZoneComponent(N_ID_ZONE) <> -1 Then
		lErrorNumber = GetZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckZoneFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.ZoneCode.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el código del registro.');" & vbNewLine
						Response.Write "oForm.ZoneCode.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ZoneName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.ZoneName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckZoneFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ZoneFrm"" ID=""ZoneFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckZoneFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Zones"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aZoneComponent(N_ID_ZONE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aZoneComponent(N_PARENT_ID_ZONE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZonePath"" ID=""ZonePathHdn"" VALUE=""" & aZoneComponent(S_PATH_ZONE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDateOld"" ID=""StartDateOldHdn"" VALUE=""" & aZoneComponent(N_START_DATE_ZONE) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ZoneCode"" ID=""ZoneCodeTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aZoneComponent(S_CODE_ZONE)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ZoneName"" ID=""ZoneNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aZoneComponent(S_NAME_ZONE)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Área geográfica:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ZoneTypeID"" ID=""ZoneTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "ZoneTypes", "ZoneTypeID", "ZoneTypeName", "(ZoneTypeID>-1)", "ZoneTypeName", aZoneComponent(N_ZONE_TYPE_ID_ZONE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Dirección:<BR /></FONT>"
					Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""255"" CLASS=""TextFields"">" & CleanStringForHTML(aZoneComponent(S_DESCRIPTION_ZONE)) & "</TEXTAREA>"
				Response.Write "</TD></TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aZoneComponent(N_START_DATE_ZONE), "Start", N_FORM_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aZoneComponent(N_END_DATE_ZONE), "End", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
							If aZoneComponent(N_ACTIVE_ZONE) = 1 Then Response.Write " CHECKED=""1"""
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
							If aZoneComponent(N_ACTIVE_ZONE) = 0 Then Response.Write " CHECKED=""0"""
						Response.Write " />No&nbsp;&nbsp;&nbsp;"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If aZoneComponent(N_ID_ZONE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveZoneWngDiv']); ZoneFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ZoneID=" & aZoneComponent(N_ID_ZONE) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveZoneWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayZoneForm = lErrorNumber
	Err.Clear
End Function

Function DisplayZoneAsHiddenFields(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a zone using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayZoneAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aZoneComponent(N_ID_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aZoneComponent(N_PARENT_ID_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneCode"" ID=""ZoneCodeHdn"" VALUE=""" & aZoneComponent(S_CODE_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneName"" ID=""ZoneNameHdn"" VALUE=""" & aZoneComponent(S_NAME_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Description"" ID=""DescriptionHdn"" VALUE=""" & aZoneComponent(S_DESCRIPTION_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZonePath"" ID=""ZonePathHdn"" VALUE=""" & aZoneComponent(S_PATH_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aZoneComponent(N_START_DATE_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aZoneComponent(N_END_DATE_ZONE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aZoneComponent(N_ACTIVE_ZONE) & """ />"

	DisplayZoneAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of a zone
'Inputs:  oRequest, oADODBConnection, aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayZonePath"
	Dim sFullPath
	Dim sTempPath
	Dim lZoneID
	Dim bFirst
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	sFullPath = ""
	bFirst = True
	lZoneID = CLng(aZoneComponent(N_PARENT_ID_ZONE))
	Do While (lZoneID <> -1)
		sErrorDescription = "No se pudo obtener la ruta de la zona."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName, ParentID, ZonePath From Zones Where (ZoneID=" & lZoneID & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If bFirst Then
					sFullPath = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value)) & "</B>" & sFullPath
					bFirst = False
				Else
					sTempPath = "<A "
						If (aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) = CLng(oRecordset.Fields("ZoneID").Value)) Or (InStr(1, CStr(oRecordset.Fields("ZonePath").Value), ("," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & ","), vbBinaryCompare) > 0) Then sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & lZoneID & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ZoneID=" & lZoneID & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
					sTempPath = sTempPath & ">" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value)) & "</A> > "
					sFullPath = sTempPath & sFullPath
				End If
				lZoneID = CLng(oRecordset.Fields("ParentID").Value)
			Else
				lZoneID = -1
			End If
		Else
			lZoneID = -1
		End If
	Loop
	Response.Write sFullPath

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayZonePath = lErrorNumber
	Err.Clear
End Function

Function DisplayZonePathAsText(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of a zone
'Inputs:  oRequest, oADODBConnection, aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayZonePathAsText"
	Dim sFullPath
	Dim sTempPath
	Dim lZoneID
	Dim bFirst
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aZoneComponent(B_COMPONENT_INITIALIZED_ZONE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeZoneComponent(oRequest, aZoneComponent)
	End If

	sFullPath = ""
	bFirst = True
	lZoneID = CLng(aZoneComponent(N_PARENT_ID_ZONE))
	Do While (lZoneID <> -1)
		sErrorDescription = "No se pudo obtener la ruta de la zona."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName, ParentID, ZonePath From Zones Where (ZoneID=" & lZoneID & ")", "ZoneComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If bFirst Then
					sFullPath = CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value)) & sFullPath
					bFirst = False
				Else
					sTempPath = CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value)) & " > "
					sFullPath = sTempPath & sFullPath
				End If
				lZoneID = CLng(oRecordset.Fields("ParentID").Value)
			Else
				lZoneID = -1
			End If
		Else
			lZoneID = -1
		End If
	Loop

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayZonePathAsText = sFullPath
	Err.Clear
End Function

Function DisplayZonesTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aZoneComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the zone  from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aZoneComponent
'Outputs: aZoneComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayZonesTable"
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
	Dim dTotalAmount
	Dim lErrorNumber

	lErrorNumber = GetZones(oRequest, oADODBConnection, aZoneComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""500"" BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (Not bForExport) Then
					asColumnsTitles = Split("&nbsp;,Código,Nombre,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,300,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Código,Nombre", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,380", ",", -1, vbBinaryCompare)
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

				dTotalAmount = 0
				asCellAlignments = Split(",,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					If (CInt(oRecordset.Fields("Active").Value) = 0) Or (CLng(oRecordset.Fields("EndDate").Value) < 30000000) Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ZoneID").Value), oRequest("ZoneID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ZoneID"" ID=""ZoneIDRd"" VALUE=""" & CStr(oRecordset.Fields("ZoneID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ZoneID"" ID=""ZoneIDChk"" VALUE=""" & CStr(oRecordset.Fields("ZoneID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;<A"
							If Not bForExport Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value)) & sBoldEnd & sFontEnd & "</A>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
							If Not bForExport Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd & "</A>"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd
					End If
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Zones&ZoneID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&Tab=1&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Zones&ZoneID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&Tab=1&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							If CInt(oRecordset.Fields("Active").Value) = 0 Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Zones&ZoneID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
							Else
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Zones&ZoneID=" & CStr(oRecordset.Fields("ZoneID").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;"

					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
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
	DisplayZonesTable = lErrorNumber
	Err.Clear
End Function
%>